<#
.SYNOPSIS
    SAFE Disk Cleanup Script for Windows 11 Professional 25H2 (Build 26200)
.DESCRIPTION
    Conservative version. Only touches caches and temp data that Windows
    regenerates automatically. NO irreversible actions by default.

    What this script DOES:
      - Cleans user/system Temp (files older than 7 days only)
      - Cleans Windows Update download cache (SoftwareDistribution\Download)
      - Cleans Delivery Optimization cache
      - Cleans browser HTTP caches (Edge, Chrome) - keeps cookies/sessions/passwords
      - Cleans Teams / Office file caches
      - Cleans Windows Error Reporting queues
      - Empties Recycle Bin (with confirmation prompt)
      - Runs DISM /AnalyzeComponentStore (READ-ONLY report, no changes)
      - Reports top 15 largest folders for manual review

    What this script does NOT do (vs. the aggressive version):
      - NO  DISM /StartComponentCleanup /ResetBase  (irreversible)
      - NO  cleanmgr /sagerun  (can hit "Previous Installations" / Windows.old)
      - NO  removal of Windows.old
      - NO  Prefetch wipe (Windows uses it to speed up app launch)
      - NO  CBS / Panther / WindowsUpdate log deletion
      - NO  crash dump (MEMORY.DMP / Minidump) deletion
      - NO  hibernation disable
      - NO  pagefile changes
      - NO  service stop/start (avoids breaking in-progress updates)

.NOTES
    Requires: Run as Administrator
    Tested on: Windows 11 Pro 25H2 (Build 26200) x64
#>

#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [switch]$DryRun,            # Show what would be deleted, change nothing
    [int]$TempAgeDays = 7,      # Only delete temp files older than N days
    [switch]$SkipRecycleBin,    # Skip emptying Recycle Bin
    [switch]$SkipBrowserCache   # Skip browser cache cleanup
)

$ErrorActionPreference = 'SilentlyContinue'
$LogFile = "$env:SystemDrive\Cleanup-Win11-Safe_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$Level] $Message"
    $color = switch ($Level) {
        'ERROR' { 'Red' }
        'WARN'  { 'Yellow' }
        'OK'    { 'Green' }
        default { 'Cyan' }
    }
    Write-Host $line -ForegroundColor $color
    Add-Content -Path $LogFile -Value $line
}

function Get-FreeSpaceGB {
    $drive = Get-PSDrive -Name ($env:SystemDrive.TrimEnd(':'))
    [math]::Round($drive.Free / 1GB, 2)
}

function Get-FolderSizeMB {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return 0 }
    try {
        $sum = (Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue |
                Measure-Object -Property Length -Sum).Sum
        return [math]::Round($sum / 1MB, 2)
    } catch { return 0 }
}

function Remove-OldFiles {
    <#
        Deletes files older than $AgeDays inside $Path.
        Keeps the folder structure intact and skips files locked by running processes.
    #>
    param(
        [string]$Path,
        [string]$Label,
        [int]$AgeDays = 7
    )
    if (-not (Test-Path $Path)) { return }

    $cutoff = (Get-Date).AddDays(-$AgeDays)
    $items = Get-ChildItem -Path $Path -Recurse -Force -File -ErrorAction SilentlyContinue |
             Where-Object { $_.LastWriteTime -lt $cutoff }

    $count = ($items | Measure-Object).Count
    $size  = ($items | Measure-Object -Property Length -Sum).Sum
    $sizeMB = [math]::Round(($size / 1MB), 2)

    if ($count -eq 0) {
        Write-Log "$Label - nothing older than $AgeDays days"
        return
    }

    if ($DryRun) {
        Write-Log "[DRYRUN] $Label - would delete $count files ($sizeMB MB) older than $AgeDays days"
        return
    }

    $removed = 0
    foreach ($f in $items) {
        try {
            Remove-Item -Path $f.FullName -Force -ErrorAction Stop
            $removed++
        } catch {
            # File locked or in use - skip silently
        }
    }
    Write-Log "$Label - removed $removed of $count files (~$sizeMB MB)" 'OK'
}

function Remove-CacheFolder {
    <#
        Empties the CONTENTS of a cache folder but keeps the folder itself.
        Skips locked files (no service stops).
    #>
    param([string]$Path, [string]$Label)
    if (-not (Test-Path $Path)) {
        Write-Log "$Label - not present, skipping"
        return
    }

    $sizeMB = Get-FolderSizeMB -Path $Path

    if ($DryRun) {
        Write-Log "[DRYRUN] $Label - would clean ~$sizeMB MB at $Path"
        return
    }

    $items = Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue
    $removed = 0; $skipped = 0
    foreach ($i in $items) {
        try {
            Remove-Item -Path $i.FullName -Recurse -Force -ErrorAction Stop
            $removed++
        } catch {
            $skipped++
        }
    }
    Write-Log "$Label - cleaned ~$sizeMB MB ($removed items removed, $skipped locked/skipped)" 'OK'
}

# ============================================================
# START
# ============================================================
Clear-Host
Write-Log "==============================================="
Write-Log " Windows 11 25H2 SAFE Disk Cleanup"
Write-Log " Host: $env:COMPUTERNAME   User: $env:USERNAME"
Write-Log " Log : $LogFile"
Write-Log " Mode: $(if ($DryRun) {'DRY RUN (no changes)'} else {'LIVE'})"
Write-Log " Temp file age threshold: $TempAgeDays days"
Write-Log "==============================================="

$freeBefore = Get-FreeSpaceGB
Write-Log "Free space BEFORE: $freeBefore GB"

# ------------------------------------------------------------
# 1. Temp folders - only files older than N days
#    (avoids breaking apps that have temp files open right now)
# ------------------------------------------------------------
Write-Log "--- Cleaning Temp folders (files older than $TempAgeDays days) ---"
Remove-OldFiles -Path "$env:WinDir\Temp"        -Label "System Temp"            -AgeDays $TempAgeDays
Remove-OldFiles -Path "$env:LOCALAPPDATA\Temp"  -Label "User Temp ($env:USERNAME)" -AgeDays $TempAgeDays

# Other user profiles
Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notin @('Public','Default','Default User','All Users') } |
    ForEach-Object {
        $userTemp = Join-Path $_.FullName "AppData\Local\Temp"
        if (Test-Path $userTemp) {
            Remove-OldFiles -Path $userTemp -Label "Temp for $($_.Name)" -AgeDays $TempAgeDays
        }
    }

# ------------------------------------------------------------
# 2. Windows Update download cache
#    Safe: Windows re-downloads what it needs. We do NOT stop wuauserv;
#    locked files are simply skipped.
# ------------------------------------------------------------
Write-Log "--- Cleaning Windows Update download cache ---"
Remove-CacheFolder -Path "$env:WinDir\SoftwareDistribution\Download" -Label "WU Download cache"

# ------------------------------------------------------------
# 3. Delivery Optimization cache (safe, peer-to-peer cache only)
# ------------------------------------------------------------
Write-Log "--- Cleaning Delivery Optimization cache ---"
Remove-CacheFolder `
    -Path "$env:WinDir\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache" `
    -Label "Delivery Optimization cache"

# ------------------------------------------------------------
# 4. Windows Error Reporting queues (safe)
# ------------------------------------------------------------
Write-Log "--- Cleaning Windows Error Reporting ---"
Remove-CacheFolder -Path "$env:ProgramData\Microsoft\Windows\WER\ReportArchive" -Label "WER ReportArchive"
Remove-CacheFolder -Path "$env:ProgramData\Microsoft\Windows\WER\ReportQueue"   -Label "WER ReportQueue"

# ------------------------------------------------------------
# 5. Browser HTTP caches (kept: cookies, history, passwords, sessions)
# ------------------------------------------------------------
if (-not $SkipBrowserCache) {
    Write-Log "--- Cleaning browser HTTP caches (sessions/passwords preserved) ---"
    $browserPaths = @(
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache";       Label = "Edge Cache" }
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Code Cache";  Label = "Edge Code Cache" }
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\GPUCache";    Label = "Edge GPUCache" }
        @{ Path = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache";        Label = "Chrome Cache" }
        @{ Path = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Code Cache";   Label = "Chrome Code Cache" }
        @{ Path = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\GPUCache";     Label = "Chrome GPUCache" }
    )
    foreach ($b in $browserPaths) {
        Remove-CacheFolder -Path $b.Path -Label $b.Label
    }
} else {
    Write-Log "Skipping browser cache (-SkipBrowserCache)"
}

# ------------------------------------------------------------
# 6. Teams & Office caches (safe; both regenerate on next start)
# ------------------------------------------------------------
Write-Log "--- Cleaning Teams / Office caches ---"
Remove-CacheFolder -Path "$env:LOCALAPPDATA\Microsoft\Teams\Cache"          -Label "Teams (classic) Cache"
Remove-CacheFolder -Path "$env:LOCALAPPDATA\Microsoft\Teams\Code Cache"     -Label "Teams (classic) Code Cache"
Remove-CacheFolder -Path "$env:LOCALAPPDATA\Microsoft\Teams\GPUCache"       -Label "Teams (classic) GPUCache"
Remove-CacheFolder -Path "$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeFileCache" -Label "Office File Cache"

# ------------------------------------------------------------
# 7. Recycle Bin - with explicit confirmation
# ------------------------------------------------------------
if (-not $SkipRecycleBin) {
    if ($DryRun) {
        Write-Log "[DRYRUN] Would prompt to empty Recycle Bin"
    } else {
        $answer = Read-Host "Empty the Recycle Bin? [y/N]"
        if ($answer -match '^(y|yes|s|si|sí)$') {
            try {
                Clear-RecycleBin -Force -ErrorAction Stop
                Write-Log "Recycle Bin emptied" 'OK'
            } catch {
                Write-Log "Recycle Bin: $($_.Exception.Message)" 'WARN'
            }
        } else {
            Write-Log "Recycle Bin skipped by user"
        }
    }
} else {
    Write-Log "Skipping Recycle Bin (-SkipRecycleBin)"
}

# ------------------------------------------------------------
# 8. DISM component store - READ-ONLY analysis
#    Tells you IF a component cleanup is recommended, without doing it.
# ------------------------------------------------------------
Write-Log "--- Analyzing component store (read-only) ---"
if (-not $DryRun) {
    $dismOutput = Dism.exe /Online /Cleanup-Image /AnalyzeComponentStore 2>&1 | Out-String
    Write-Log $dismOutput
    Write-Log "If DISM reports 'Component Store Cleanup Recommended : Yes', you can run" 'WARN'
    Write-Log "manually:  Dism.exe /Online /Cleanup-Image /StartComponentCleanup" 'WARN'
    Write-Log "(without /ResetBase, so update rollback stays possible)" 'WARN'
}

# ------------------------------------------------------------
# 9. Largest folders report (top 15)
# ------------------------------------------------------------
Write-Log "--- Scanning top 15 largest folders under C:\ ---"
$top = Get-ChildItem -Path "$env:SystemDrive\" -Directory -Force -ErrorAction SilentlyContinue |
    ForEach-Object {
        $size = (Get-ChildItem $_.FullName -Recurse -Force -ErrorAction SilentlyContinue |
                 Measure-Object -Property Length -Sum).Sum
        [pscustomobject]@{ Path = $_.FullName; SizeGB = [math]::Round($size/1GB,2) }
    } | Sort-Object SizeGB -Descending | Select-Object -First 15

$top | Format-Table -AutoSize | Out-String | ForEach-Object { Write-Log $_ }

# ------------------------------------------------------------
# FINAL REPORT
# ------------------------------------------------------------
$freeAfter  = Get-FreeSpaceGB
$reclaimed  = [math]::Round($freeAfter - $freeBefore, 2)

Write-Log "==============================================="
Write-Log " SAFE CLEANUP COMPLETE"
Write-Log " Free space BEFORE : $freeBefore GB"
Write-Log " Free space AFTER  : $freeAfter GB"
Write-Log " Reclaimed         : $reclaimed GB"
Write-Log " Log saved to      : $LogFile"
Write-Log "==============================================="

Write-Log ""
Write-Log "Next steps if more space is needed (manual, your call):" 'WARN'
Write-Log "  1) Settings > System > Storage > Cleanup recommendations" 'WARN'
Write-Log "  2) Settings > Apps > Installed apps - sort by size, uninstall unused" 'WARN'
Write-Log "  3) Move OneDrive / Documents to another drive if available" 'WARN'
Write-Log "  4) Review the top-15 folder list above for unexpected large folders" 'WARN'
