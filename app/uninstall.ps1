# uninstall.ps1 -- EntryPass-MiHCM Sync Uninstaller
# Removes the Windows Task Scheduler task created by setup.ps1.
# Does NOT delete config.json, sync_log.txt, or any data files.
#
# Requirements: PowerShell 5.1+, Windows 10/11

$ErrorActionPreference = "Continue"

# ============================================================
# SCRIPT DIRECTORY
# ============================================================
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($scriptDir)) {
    $scriptDir = Split-Path -Parent $PSCommandPath
}
if ([string]::IsNullOrWhiteSpace($scriptDir)) {
    $scriptDir = Get-Location
}

Write-Host "EntryPass-MiHCM Sync -- Uninstaller"
Write-Host "===================================="
Write-Host ""

# ============================================================
# READ CONFIG to get location code for task name
# ============================================================
$configFile = Join-Path $scriptDir "config.json"
$locationCode = $null

if (Test-Path $configFile) {
    try {
        $config = Get-Content $configFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $locationCode = $config.location
    } catch {
        Write-Host "WARNING: Could not read config.json -- $_"
    }
}

if ([string]::IsNullOrWhiteSpace($locationCode)) {
    Write-Host "Could not determine Location Code from config.json."
    $locationCode = Read-Host "Enter the Location Code used during setup (e.g. PEMO)"
    $locationCode = $locationCode.Trim().ToUpper()
}

$taskName = "EntryPass-MiHCM Sync - $locationCode"
Write-Host "Task to remove : $taskName"
Write-Host ""

# ============================================================
# CHECK IF TASK EXISTS
# ============================================================
$taskExists = $false
try {
    $check = & schtasks.exe /Query /TN $taskName 2>&1
    if ($LASTEXITCODE -eq 0) {
        $taskExists = $true
    }
} catch {}

if (-not $taskExists) {
    Write-Host "Scheduled task '$taskName' was not found."
    Write-Host "It may have already been removed, or was never installed."
    Write-Host ""
    Read-Host "Press Enter to close..."
    exit 0
}

# ============================================================
# CONFIRM BEFORE REMOVING
# ============================================================
Write-Host "This will remove the scheduled task: $taskName"
Write-Host "Your config.json and log files will NOT be deleted."
Write-Host ""
$confirm = Read-Host "Type YES to confirm removal"

if ($confirm.Trim().ToUpper() -ne "YES") {
    Write-Host ""
    Write-Host "Uninstall cancelled. No changes made."
    Write-Host ""
    Read-Host "Press Enter to close..."
    exit 0
}

# ============================================================
# REMOVE SCHEDULED TASK
# ============================================================
Write-Host ""
Write-Host "Removing task '$taskName'..."

try {
    $result = & schtasks.exe /Delete /TN $taskName /F 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Task removed successfully."
        Write-Host ""
        Write-Host "Note: config.json and sync_log.txt have NOT been deleted."
        Write-Host "You may delete this folder manually if no longer needed."
    } else {
        Write-Host "ERROR: schtasks returned exit code $LASTEXITCODE"
        Write-Host "Output: $result"
        Write-Host ""
        Write-Host "You can remove the task manually from Task Scheduler."
    }
} catch {
    Write-Host "ERROR: $_"
    Write-Host "You can remove the task manually from Task Scheduler."
}

Write-Host ""
Read-Host "Press Enter to close..."
