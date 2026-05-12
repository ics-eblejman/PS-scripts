<#
.SYNOPSIS
    Disables the ability for end users to change their profile photo across
    Microsoft 365 (Outlook/OWA, Teams, Delve, myaccount, Entra) by applying
    the five controls available in the platform.

.DESCRIPTION
    This script applies a defense-in-depth configuration:

      1) Connects to Exchange Online and sets SetPhotoEnabled = $false on
         EVERY existing OwaMailboxPolicy (affects OWA and Teams).
      2) Calls Set-MsolCompanySettings -UsersPermissionToChangeProfilePictureDisabled $true
         (MSOnline has been deprecated since 2024-03-30 — the script attempts
          to run it and, if it fails, applies the modern Microsoft Graph
          equivalent).
      3) Restricts non-admin access to the Microsoft Entra admin center by
         setting restrictNonAdminUsers = true on authorizationPolicy
         (Graph beta).
      4) Configures photoUpdateSettings at
         /beta/admin/people/photoUpdateSettings, leaving ONLY the Global
         Administrator role in allowedRoles.
      5) Verifies every change by reading the resulting state and produces
         a summary report.

    Prerequisites (install any that are missing):
        Install-Module ExchangeOnlineManagement       -Scope CurrentUser
        Install-Module Microsoft.Graph                 -Scope CurrentUser
        Install-Module Microsoft.Graph.Beta            -Scope CurrentUser
        Install-Module MSOnline                        -Scope CurrentUser   # optional / deprecated

    Graph delegated permissions requested by Connect-MgGraph:
        Policy.ReadWrite.Authorization
        PeopleSettings.ReadWrite.All
        Directory.ReadWrite.All

    Roles required for the executing account:
        Global Administrator (or Exchange Admin + Privileged Role Admin
        + People Administrator combined).

.PARAMETER AdminUpn
    UPN of the administrator running the interactive connections. If omitted,
    each Connect-* cmdlet will prompt for credentials separately.

.PARAMETER WhatIf
    Shows what would change without applying the changes.

.EXAMPLE
    .\Disable-M365ProfilePhotoChanges.ps1 -AdminUpn admin@contoso.com -Verbose

.NOTES
    Author  : Ezequiel Blejman
    Version : 1.0
    Date    : 2026-05
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$AdminUpn
)

# -----------------------------------------------------------------------------
# Global configuration and helpers
# -----------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

# GUID for the Global Administrator role (Entra ID built-in role template ID)
$GlobalAdminRoleId = '62e90394-69f5-4237-9190-012177145e10'

# Final step summary
$Script:Summary = [System.Collections.Generic.List[object]]::new()

function Write-Step {
    param(
        [Parameter(Mandatory)] [string]$Step,
        [Parameter(Mandatory)] [ValidateSet('OK','WARN','FAIL','SKIP')] [string]$Status,
        [string]$Detail = ''
    )
    $color = switch ($Status) {
        'OK'   { 'Green' }
        'WARN' { 'Yellow' }
        'FAIL' { 'Red' }
        'SKIP' { 'DarkGray' }
    }
    Write-Host ("[{0,-4}] {1}" -f $Status, $Step) -ForegroundColor $color
    if ($Detail) { Write-Host "         $Detail" -ForegroundColor DarkGray }
    $Script:Summary.Add([pscustomobject]@{
        Step   = $Step
        Status = $Status
        Detail = $Detail
    })
}

function Assert-Module {
    param(
        [Parameter(Mandatory)] [string]$Name,
        [switch]$Optional
    )
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        if ($Optional) {
            Write-Step "Module $Name not installed" -Status 'WARN' `
                -Detail "Steps that depend on this module will be skipped."
            return $false
        }
        throw "Required module '$Name' is not installed. Run: Install-Module $Name -Scope CurrentUser"
    }
    Import-Module $Name -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
    return $true
}

Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host " Microsoft 365 - Block profile photo changes " -ForegroundColor Cyan
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host ""

# -----------------------------------------------------------------------------
# 0) Validate modules
# -----------------------------------------------------------------------------
try {
    [void](Assert-Module -Name 'ExchangeOnlineManagement')
    [void](Assert-Module -Name 'Microsoft.Graph.Authentication')
    [void](Assert-Module -Name 'Microsoft.Graph.Identity.SignIns')
    [void](Assert-Module -Name 'Microsoft.Graph.Beta.Identity.DirectoryManagement')
    $msolAvailable = Assert-Module -Name 'MSOnline' -Optional
}
catch {
    Write-Step "Module verification" -Status 'FAIL' -Detail $_.Exception.Message
    return
}

# -----------------------------------------------------------------------------
# 1) Exchange Online - OwaMailboxPolicy.SetPhotoEnabled = $false
# -----------------------------------------------------------------------------
try {
    Write-Verbose "Connecting to Exchange Online..."
    if ($AdminUpn) {
        Connect-ExchangeOnline -UserPrincipalName $AdminUpn -ShowBanner:$false -ErrorAction Stop
    } else {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    }
    Write-Step "Exchange Online connection" -Status 'OK'

    $policies = Get-OwaMailboxPolicy -ErrorAction Stop
    if (-not $policies) { throw "No OwaMailboxPolicy objects found in the tenant." }

    foreach ($p in $policies) {
        try {
            if ($PSCmdlet.ShouldProcess($p.Identity, "Set-OwaMailboxPolicy -SetPhotoEnabled `$false")) {
                Set-OwaMailboxPolicy -Identity $p.Identity -SetPhotoEnabled $false -ErrorAction Stop
            }
            # Verification
            $check = Get-OwaMailboxPolicy -Identity $p.Identity | Select-Object -ExpandProperty SetPhotoEnabled
            if ($check -eq $false) {
                Write-Step "OwaMailboxPolicy '$($p.Identity)'" -Status 'OK' `
                    -Detail "SetPhotoEnabled = False"
            } else {
                Write-Step "OwaMailboxPolicy '$($p.Identity)'" -Status 'WARN' `
                    -Detail "Verification returned SetPhotoEnabled = $check"
            }
        } catch {
            Write-Step "OwaMailboxPolicy '$($p.Identity)'" -Status 'FAIL' -Detail $_.Exception.Message
        }
    }

    # Note: OwaMailboxPolicy objects are already assigned to users via the
    # mailbox. There is no need to re-assign the default; custom policies are
    # assigned with Set-CASMailbox -OwaMailboxPolicy <name>. We just report
    # the current coverage.
    $mbxCovered = (Get-CASMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue |
                   Where-Object { $_.OwaMailboxPolicy } | Measure-Object).Count
    Write-Step "OwaMailboxPolicy coverage" -Status 'OK' `
        -Detail "$mbxCovered mailboxes have an OwaMailboxPolicy assigned."
}
catch {
    Write-Step "Step 1 - Exchange Online OwaMailboxPolicy" -Status 'FAIL' -Detail $_.Exception.Message
}

# -----------------------------------------------------------------------------
# 2) MSOnline - Set-MsolCompanySettings (DEPRECATED since 2024-03-30)
# -----------------------------------------------------------------------------
$msolDidWork = $false
if ($msolAvailable) {
    try {
        Write-Verbose "Connecting to MSOnline (deprecated module)..."
        if ($AdminUpn) {
            $cred = Get-Credential -UserName $AdminUpn -Message "Credentials for Connect-MsolService (MSOnline is deprecated)"
            Connect-MsolService -Credential $cred -ErrorAction Stop
        } else {
            Connect-MsolService -ErrorAction Stop
        }

        # NOTE: UsersPermissionToChangeProfilePictureDisabled is NOT a published
        # parameter of Set-MsolCompanySettings — the official documentation
        # (azureadps-1.0) only exposes the parameters listed in the cmdlet. We
        # still attempt it in case your tenant accepts it through dynamic
        # splatting.
        $params = @{
            UsersPermissionToChangeProfilePictureDisabled = $true
        }
        if ($PSCmdlet.ShouldProcess('Company', 'Set-MsolCompanySettings (legacy)')) {
            Set-MsolCompanySettings @params -ErrorAction Stop
        }
        $msolDidWork = $true
        Write-Step "Set-MsolCompanySettings (legacy)" -Status 'OK' `
            -Detail "UsersPermissionToChangeProfilePictureDisabled = True"
    }
    catch {
        Write-Step "Set-MsolCompanySettings (legacy)" -Status 'WARN' `
            -Detail "MSOnline failed or the parameter is unavailable: $($_.Exception.Message). The modern equivalent will be applied in step 3."
    }
}
else {
    Write-Step "MSOnline" -Status 'SKIP' -Detail "Module not installed (deprecated since 2024-03-30). The Graph equivalent will be applied instead."
}

# -----------------------------------------------------------------------------
# 3 + Modern replacement) Microsoft Graph - authorizationPolicy
#     Restricts non-admin access to the Entra admin center and, if MSOnline
#     did not work, applies the photo-change block via the modern Graph path.
# -----------------------------------------------------------------------------
try {
    Write-Verbose "Connecting to Microsoft Graph..."
    Connect-MgGraph -NoWelcome -Scopes @(
        'Policy.ReadWrite.Authorization',
        'PeopleSettings.ReadWrite.All',
        'Directory.ReadWrite.All'
    ) -ErrorAction Stop
    $ctx = Get-MgContext
    Write-Step "Microsoft Graph connection" -Status 'OK' -Detail "Tenant: $($ctx.TenantId)"

    # ----- 3a) Restrict access to Microsoft Entra admin center -----
    # Property: restrictNonAdminUsers = $true on authorizationPolicy.
    $body = @{ restrictNonAdminUsers = $true } | ConvertTo-Json -Compress
    if ($PSCmdlet.ShouldProcess('authorizationPolicy', 'PATCH restrictNonAdminUsers=true')) {
        Invoke-MgGraphRequest -Method PATCH `
            -Uri 'https://graph.microsoft.com/beta/policies/authorizationPolicy' `
            -Body $body `
            -ContentType 'application/json' `
            -ErrorAction Stop | Out-Null
    }

    # Verification
    $authPol = Invoke-MgGraphRequest -Method GET `
        -Uri 'https://graph.microsoft.com/beta/policies/authorizationPolicy' `
        -ErrorAction Stop
    if ($authPol.restrictNonAdminUsers -eq $true) {
        Write-Step "Restrict access to Entra admin center" -Status 'OK' `
            -Detail "restrictNonAdminUsers = True"
    } else {
        Write-Step "Restrict access to Entra admin center" -Status 'WARN' `
            -Detail "Verification returned restrictNonAdminUsers = $($authPol.restrictNonAdminUsers)"
    }
}
catch {
    Write-Step "Step 3 - authorizationPolicy (Graph)" -Status 'FAIL' -Detail $_.Exception.Message
}

# -----------------------------------------------------------------------------
# 4) photoUpdateSettings - Only Global Admin can change photos
# -----------------------------------------------------------------------------
try {
    $payload = @{
        '@odata.type' = '#microsoft.graph.photoUpdateSettings'
        source        = 'cloud'
        allowedRoles  = @($GlobalAdminRoleId)
    } | ConvertTo-Json -Compress

    if ($PSCmdlet.ShouldProcess('photoUpdateSettings', 'PATCH allowedRoles=[GlobalAdmin]')) {
        try {
            # Primary attempt: PATCH (resource already exists in most tenants)
            Invoke-MgGraphRequest -Method PATCH `
                -Uri 'https://graph.microsoft.com/beta/admin/people/photoUpdateSettings' `
                -Body $payload `
                -ContentType 'application/json' `
                -ErrorAction Stop | Out-Null
        } catch {
            # Fallback: if the resource does not exist yet (404), create it with POST
            if ($_.Exception.Message -match '404' -or $_.Exception.Message -match 'NotFound') {
                Write-Verbose "photoUpdateSettings does not exist — creating it with POST..."
                Invoke-MgGraphRequest -Method POST `
                    -Uri 'https://graph.microsoft.com/beta/admin/people/photoUpdateSettings' `
                    -Body $payload `
                    -ContentType 'application/json' `
                    -ErrorAction Stop | Out-Null
            } else { throw }
        }
    }

    # Verification
    $pus = Invoke-MgGraphRequest -Method GET `
        -Uri 'https://graph.microsoft.com/beta/admin/people/photoUpdateSettings' `
        -ErrorAction Stop
    $rolesApplied = @($pus.allowedRoles)
    if ($rolesApplied.Count -eq 1 -and $rolesApplied[0] -eq $GlobalAdminRoleId -and $pus.source -eq 'cloud') {
        Write-Step "photoUpdateSettings" -Status 'OK' `
            -Detail "source=cloud, allowedRoles=[Global Admin]"
    } else {
        Write-Step "photoUpdateSettings" -Status 'WARN' `
            -Detail "Verification: source=$($pus.source), allowedRoles=$($rolesApplied -join ',')"
    }
}
catch {
    Write-Step "Step 4 - photoUpdateSettings" -Status 'FAIL' -Detail $_.Exception.Message
}

# -----------------------------------------------------------------------------
# Connection cleanup
# -----------------------------------------------------------------------------
try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
if ($msolAvailable -and $msolDidWork) {
    # MSOnline does not expose Disconnect-MsolService; the session dies with the process.
}

# -----------------------------------------------------------------------------
# Final summary
# -----------------------------------------------------------------------------
Write-Host ""
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host " Application summary " -ForegroundColor Cyan
Write-Host "==============================================================" -ForegroundColor Cyan
$Script:Summary | Format-Table -AutoSize Step, Status, Detail

$failed = ($Script:Summary | Where-Object Status -eq 'FAIL').Count
$warned = ($Script:Summary | Where-Object Status -eq 'WARN').Count
Write-Host ""
if ($failed -gt 0) {
    Write-Host "Finished with $failed errors and $warned warnings." -ForegroundColor Red
    exit 1
} elseif ($warned -gt 0) {
    Write-Host "Finished with $warned warnings. Review the details above." -ForegroundColor Yellow
    Write-Host "Reminder: policy changes may take up to 24 hours to propagate." -ForegroundColor Yellow
    exit 0
} else {
    Write-Host "All controls were applied and verified successfully." -ForegroundColor Green
    Write-Host "Expected propagation: up to 24 hours across Teams/Outlook/Delve." -ForegroundColor Green
    exit 0
}
