# setup.ps1 -- EntryPass-MiHCM Sync Setup Wizard
# GUI configuration tool. Run once to create config.json and optionally
# install a Windows Task Scheduler task for automated sync.
#
# Requirements: PowerShell 5.1+, Windows 10/11 (WinForms)

$ErrorActionPreference = "Stop"

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

# ============================================================
# LOAD WINFORMS
# ============================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ============================================================
# LOAD EXISTING CONFIG (if present) for pre-filling fields
# ============================================================
$existingConfig = $null
$configFile = Join-Path $scriptDir "config.json"
if (Test-Path $configFile) {
    try {
        $existingConfig = Get-Content $configFile -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        # Ignore parse errors -- start fresh
    }
}

function Get-ConfigValue {
    param([string]$Field, [string]$Default = "")
    if ($existingConfig -and $existingConfig.$Field) {
        return $existingConfig.$Field
    }
    return $Default
}

# ============================================================
# BUILD THE FORM
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text            = "EntryPass-MiHCM Sync - Setup"
$form.Size            = New-Object System.Drawing.Size(520, 680)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox     = $false
$form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)
$form.BackColor       = [System.Drawing.Color]::White

# -- Branding header --
$lblBrand = New-Object System.Windows.Forms.Label
$lblBrand.Text      = "Dajayana Trading"
$lblBrand.Font      = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$lblBrand.ForeColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
$lblBrand.Location  = New-Object System.Drawing.Point(16, 12)
$lblBrand.Size      = New-Object System.Drawing.Size(480, 30)
$form.Controls.Add($lblBrand)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text      = "EntryPass to MiHCM Sync -- Setup Wizard"
$lblSub.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$lblSub.ForeColor = [System.Drawing.Color]::FromArgb(80, 80, 80)
$lblSub.Location  = New-Object System.Drawing.Point(16, 44)
$lblSub.Size      = New-Object System.Drawing.Size(480, 20)
$form.Controls.Add($lblSub)

$separator = New-Object System.Windows.Forms.Label
$separator.BorderStyle = "Fixed3D"
$separator.Location    = New-Object System.Drawing.Point(16, 68)
$separator.Size        = New-Object System.Drawing.Size(480, 2)
$form.Controls.Add($separator)

# ============================================================
# HELPER -- add a label + input pair
# ============================================================
$yPos = 80

function Add-Field {
    param(
        [string]$LabelText,
        [int]$Y,
        [System.Windows.Forms.Control]$Control
    )
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text     = $LabelText
    $lbl.Location = New-Object System.Drawing.Point(16, ($Y + 3))
    $lbl.Size     = New-Object System.Drawing.Size(160, 20)
    $form.Controls.Add($lbl)
    $Control.Location = New-Object System.Drawing.Point(180, $Y)
    $Control.Size     = New-Object System.Drawing.Size(316, 24)
    $form.Controls.Add($Control)
}

# -- License Key --
$txtLicense = New-Object System.Windows.Forms.TextBox
$txtLicense.Text = Get-ConfigValue "licenseKey"
Add-Field "License Key *" $yPos $txtLicense
$yPos += 32

# -- Primary Key --
$txtPrimary = New-Object System.Windows.Forms.TextBox
$txtPrimary.Text = Get-ConfigValue "primaryKey"
Add-Field "MiHCM Primary Key *" $yPos $txtPrimary
$yPos += 32

# -- Secret Key --
$txtSecret = New-Object System.Windows.Forms.TextBox
$txtSecret.Text         = Get-ConfigValue "secretKey"
$txtSecret.PasswordChar = '*'
Add-Field "MiHCM Secret Key *" $yPos $txtSecret
$yPos += 32

# -- API Endpoint dropdown --
$cmbEndpoint = New-Object System.Windows.Forms.ComboBox
$cmbEndpoint.DropDownStyle = "DropDownList"
[void]$cmbEndpoint.Items.Add("Production (api.mihcm.com)")
[void]$cmbEndpoint.Items.Add("UAT (api.mihcm.com/uat)")
$savedEndpoint = Get-ConfigValue "apiEndpoint" "https://api.mihcm.com"
if ($savedEndpoint -match "uat") {
    $cmbEndpoint.SelectedIndex = 1
} else {
    $cmbEndpoint.SelectedIndex = 0
}
Add-Field "API Endpoint *" $yPos $cmbEndpoint
$yPos += 32

# -- Client Name --
$txtClient = New-Object System.Windows.Forms.TextBox
$txtClient.Text = Get-ConfigValue "clientName"
Add-Field "Client Name *" $yPos $txtClient
$yPos += 32

# -- Location Code --
$txtLocation = New-Object System.Windows.Forms.TextBox
$txtLocation.Text = Get-ConfigValue "location"
Add-Field "Location Code *" $yPos $txtLocation
$yPos += 32

# -- Location Description --
$txtLocationDesc = New-Object System.Windows.Forms.TextBox
$txtLocationDesc.Text = Get-ConfigValue "locationDesc"
Add-Field "Location Description" $yPos $txtLocationDesc
$yPos += 32

# -- Source Folder (text + browse button) --
$lblSourceFolder = New-Object System.Windows.Forms.Label
$lblSourceFolder.Text     = "Source Folder *"
$lblSourceFolder.Location = New-Object System.Drawing.Point(16, ($yPos + 3))
$lblSourceFolder.Size     = New-Object System.Drawing.Size(160, 20)
$form.Controls.Add($lblSourceFolder)

$txtSourceFolder = New-Object System.Windows.Forms.TextBox
$txtSourceFolder.Text     = Get-ConfigValue "sourceFolder"
$txtSourceFolder.Location = New-Object System.Drawing.Point(180, $yPos)
$txtSourceFolder.Size     = New-Object System.Drawing.Size(226, 24)
$form.Controls.Add($txtSourceFolder)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text     = "Browse..."
$btnBrowse.Location = New-Object System.Drawing.Point(412, $yPos)
$btnBrowse.Size     = New-Object System.Drawing.Size(84, 24)
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description         = "Select the folder containing EntryPass DATA*.txt files"
    $dlg.ShowNewFolderButton = $true
    if ($txtSourceFolder.Text -and (Test-Path $txtSourceFolder.Text)) {
        $dlg.SelectedPath = $txtSourceFolder.Text
    }
    if ($dlg.ShowDialog() -eq "OK") {
        $txtSourceFolder.Text = $dlg.SelectedPath
    }
})
$form.Controls.Add($btnBrowse)
$yPos += 36

# -- Separator --
$sep2 = New-Object System.Windows.Forms.Label
$sep2.BorderStyle = "Fixed3D"
$sep2.Location    = New-Object System.Drawing.Point(16, $yPos)
$sep2.Size        = New-Object System.Drawing.Size(480, 2)
$form.Controls.Add($sep2)
$yPos += 10

# -- Enable Scheduled Sync checkbox --
$chkSchedule = New-Object System.Windows.Forms.CheckBox
$chkSchedule.Text     = "Enable Scheduled Sync (Windows Task Scheduler)"
$chkSchedule.Location = New-Object System.Drawing.Point(16, $yPos)
$chkSchedule.Size     = New-Object System.Drawing.Size(480, 22)
$chkSchedule.Checked  = $false
$form.Controls.Add($chkSchedule)
$yPos += 28

# -- Schedule Frequency dropdown --
$lblFreq = New-Object System.Windows.Forms.Label
$lblFreq.Text     = "Schedule Frequency"
$lblFreq.Location = New-Object System.Drawing.Point(16, ($yPos + 3))
$lblFreq.Size     = New-Object System.Drawing.Size(160, 20)
$form.Controls.Add($lblFreq)

$cmbFrequency = New-Object System.Windows.Forms.ComboBox
$cmbFrequency.DropDownStyle = "DropDownList"
[void]$cmbFrequency.Items.Add("Every 15 minutes")
[void]$cmbFrequency.Items.Add("Every 30 minutes")
[void]$cmbFrequency.Items.Add("Hourly")
[void]$cmbFrequency.Items.Add("Every 2 hours")
[void]$cmbFrequency.Items.Add("Twice daily (8am, 6pm)")
[void]$cmbFrequency.Items.Add("Daily (6pm)")
$cmbFrequency.SelectedIndex = 1
$cmbFrequency.Location      = New-Object System.Drawing.Point(180, $yPos)
$cmbFrequency.Size          = New-Object System.Drawing.Size(316, 24)
$form.Controls.Add($cmbFrequency)
$yPos += 36

# -- Separator --
$sep3 = New-Object System.Windows.Forms.Label
$sep3.BorderStyle = "Fixed3D"
$sep3.Location    = New-Object System.Drawing.Point(16, $yPos)
$sep3.Size        = New-Object System.Drawing.Size(480, 2)
$form.Controls.Add($sep3)
$yPos += 12

# ============================================================
# ACTION BUTTONS -- row 1: Test Connection | Validate License
# ============================================================
$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text     = "Test Connection"
$btnTest.Location = New-Object System.Drawing.Point(16, $yPos)
$btnTest.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnTest)

$btnValidate = New-Object System.Windows.Forms.Button
$btnValidate.Text     = "Validate License"
$btnValidate.Location = New-Object System.Drawing.Point(174, $yPos)
$btnValidate.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnValidate)
$yPos += 40

# -- row 2: Save & Install | Save Only --
$btnSaveInstall = New-Object System.Windows.Forms.Button
$btnSaveInstall.Text      = "Save && Install"
$btnSaveInstall.Location  = New-Object System.Drawing.Point(16, $yPos)
$btnSaveInstall.Size      = New-Object System.Drawing.Size(150, 32)
$btnSaveInstall.BackColor = [System.Drawing.Color]::FromArgb(0, 102, 204)
$btnSaveInstall.ForeColor = [System.Drawing.Color]::White
$btnSaveInstall.FlatStyle = "Flat"
$form.Controls.Add($btnSaveInstall)

$btnSaveOnly = New-Object System.Windows.Forms.Button
$btnSaveOnly.Text     = "Save Only"
$btnSaveOnly.Location = New-Object System.Drawing.Point(174, $yPos)
$btnSaveOnly.Size     = New-Object System.Drawing.Size(150, 32)
$form.Controls.Add($btnSaveOnly)
$yPos += 42

# -- Status label --
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text      = ""
$lblStatus.Location  = New-Object System.Drawing.Point(16, $yPos)
$lblStatus.Size      = New-Object System.Drawing.Size(480, 40)
$lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
$form.Controls.Add($lblStatus)

# ============================================================
# HELPER -- get API endpoint URL from dropdown selection
# ============================================================
function Get-EndpointUrl {
    if ($cmbEndpoint.SelectedIndex -eq 1) {
        return "https://api.mihcm.com/uat"
    }
    return "https://api.mihcm.com"
}

# ============================================================
# HELPER -- validate required fields, return $true if OK
# ============================================================
function Validate-Fields {
    if ([string]::IsNullOrWhiteSpace($txtLicense.Text)) {
        [System.Windows.Forms.MessageBox]::Show("License Key is required.", "Validation Error", "OK", "Warning")
        return $false
    }
    if ([string]::IsNullOrWhiteSpace($txtPrimary.Text)) {
        [System.Windows.Forms.MessageBox]::Show("MiHCM Primary Key is required.", "Validation Error", "OK", "Warning")
        return $false
    }
    if ([string]::IsNullOrWhiteSpace($txtSecret.Text)) {
        [System.Windows.Forms.MessageBox]::Show("MiHCM Secret Key is required.", "Validation Error", "OK", "Warning")
        return $false
    }
    if ([string]::IsNullOrWhiteSpace($txtClient.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Client Name is required.", "Validation Error", "OK", "Warning")
        return $false
    }
    if ([string]::IsNullOrWhiteSpace($txtLocation.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Location Code is required.", "Validation Error", "OK", "Warning")
        return $false
    }
    if ([string]::IsNullOrWhiteSpace($txtSourceFolder.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Source Folder is required.", "Validation Error", "OK", "Warning")
        return $false
    }
    return $true
}

# ============================================================
# HELPER -- build and write config.json
# ============================================================
function Save-Config {
    $cfg = [ordered]@{
        licenseKey   = $txtLicense.Text.Trim()
        primaryKey   = $txtPrimary.Text.Trim()
        secretKey    = $txtSecret.Text.Trim()
        apiEndpoint  = Get-EndpointUrl
        clientName   = $txtClient.Text.Trim()
        location     = $txtLocation.Text.Trim().ToUpper()
        locationDesc = $txtLocationDesc.Text.Trim()
        sourceFolder = $txtSourceFolder.Text.Trim()
        batchSize    = 80
    }
    $json = $cfg | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($configFile, $json, [System.Text.Encoding]::UTF8)
    return $cfg
}

# ============================================================
# HELPER -- create Task Scheduler task
# ============================================================
function Install-ScheduledTask {
    param([string]$LocationCode, [string]$Frequency)

    $taskName  = "EntryPass-MiHCM Sync - $LocationCode"
    $syncScript = Join-Path $scriptDir "sync.ps1"
    $action    = "powershell.exe"
    $args      = "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File `"$syncScript`""

    # Build trigger XML based on frequency
    $triggerXml = ""
    switch ($Frequency) {
        "Every 15 minutes" {
            $triggerXml = @"
<CalendarTrigger>
  <Repetition>
    <Interval>PT15M</Interval>
    <Duration>P1D</Duration>
    <StopAtDurationEnd>false</StopAtDurationEnd>
  </Repetition>
  <StartBoundary>2024-01-01T06:00:00</StartBoundary>
</CalendarTrigger>
"@
        }
        "Every 30 minutes" {
            $triggerXml = @"
<CalendarTrigger>
  <Repetition>
    <Interval>PT30M</Interval>
    <Duration>P1D</Duration>
    <StopAtDurationEnd>false</StopAtDurationEnd>
  </Repetition>
  <StartBoundary>2024-01-01T06:00:00</StartBoundary>
</CalendarTrigger>
"@
        }
        "Hourly" {
            $triggerXml = @"
<CalendarTrigger>
  <Repetition>
    <Interval>PT1H</Interval>
    <Duration>P1D</Duration>
    <StopAtDurationEnd>false</StopAtDurationEnd>
  </Repetition>
  <StartBoundary>2024-01-01T06:00:00</StartBoundary>
</CalendarTrigger>
"@
        }
        "Every 2 hours" {
            $triggerXml = @"
<CalendarTrigger>
  <Repetition>
    <Interval>PT2H</Interval>
    <Duration>P1D</Duration>
    <StopAtDurationEnd>false</StopAtDurationEnd>
  </Repetition>
  <StartBoundary>2024-01-01T06:00:00</StartBoundary>
</CalendarTrigger>
"@
        }
        "Twice daily (8am, 6pm)" {
            # Create two separate time triggers
            $triggerXml = @"
<CalendarTrigger>
  <StartBoundary>2024-01-01T08:00:00</StartBoundary>
  <ScheduleByDay><DaysInterval>1</DaysInterval></ScheduleByDay>
</CalendarTrigger>
<CalendarTrigger>
  <StartBoundary>2024-01-01T18:00:00</StartBoundary>
  <ScheduleByDay><DaysInterval>1</DaysInterval></ScheduleByDay>
</CalendarTrigger>
"@
        }
        default {
            # Daily (6pm)
            $triggerXml = @"
<CalendarTrigger>
  <StartBoundary>2024-01-01T18:00:00</StartBoundary>
  <ScheduleByDay><DaysInterval>1</DaysInterval></ScheduleByDay>
</CalendarTrigger>
"@
        }
    }

    # Build full task XML
    $taskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>EntryPass to MiHCM attendance sync for $LocationCode</Description>
  </RegistrationInfo>
  <Triggers>
    $triggerXml
  </Triggers>
  <Principals>
    <Principal id="Author">
      <LogonType>S4U</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions>
    <Exec>
      <Command>$action</Command>
      <Arguments>$args</Arguments>
      <WorkingDirectory>$scriptDir</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@

    # Write temp XML and register via schtasks
    $tmpXml = Join-Path $env:TEMP "entrypass_sync_task.xml"
    [System.IO.File]::WriteAllText($tmpXml, $taskXml, [System.Text.Encoding]::Unicode)

    try {
        $result = & schtasks.exe /Create /TN $taskName /XML $tmpXml /F 2>&1
        Remove-Item $tmpXml -Force -ErrorAction SilentlyContinue
        if ($LASTEXITCODE -eq 0) {
            return @{ Success = $true; Message = "Task '$taskName' created successfully." }
        } else {
            return @{ Success = $false; Message = "schtasks returned exit code $LASTEXITCODE. Output: $result" }
        }
    } catch {
        Remove-Item $tmpXml -Force -ErrorAction SilentlyContinue
        return @{ Success = $false; Message = "Failed to create task: $_" }
    }
}

# ============================================================
# BUTTON -- Test Connection
# ============================================================
$btnTest.Add_Click({
    $lblStatus.Text      = "Testing connection..."
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
    $form.Refresh()

    $baseUrl   = Get-EndpointUrl
    $primary   = $txtPrimary.Text.Trim()
    $secret    = $txtSecret.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($primary) -or [string]::IsNullOrWhiteSpace($secret)) {
        $lblStatus.Text      = "ERROR: Enter Primary Key and Secret Key first."
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        return
    }

    try {
        $tokenUrl = "$baseUrl/oauth2/token?grantType=client_credentials&clientId=$primary&clientSecret=$secret"
        $raw = Invoke-WebRequest -Uri $tokenUrl -Method GET -Headers @{
            "Ocp-Apim-Subscription-Key" = $primary
        } -UseBasicParsing -TimeoutSec 15
        $resp = $raw.Content | ConvertFrom-Json
        if ($resp.accessToken) {
            $lblStatus.Text      = "Connection OK -- Token obtained. Endpoint is reachable."
            $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
        } else {
            $lblStatus.Text      = "Connected but no accessToken returned. Check keys."
            $lblStatus.ForeColor = [System.Drawing.Color]::DarkOrange
        }
    } catch {
        $statusCode = ""
        if ($_.Exception.Response) {
            $statusCode = " (HTTP $([int]$_.Exception.Response.StatusCode))"
        }
        $lblStatus.Text      = "Connection FAILED$statusCode -- $($_.Exception.Message)"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
    }
})

# ============================================================
# BUTTON -- Validate License
# ============================================================
$btnValidate.Add_Click({
    $lblStatus.Text      = "Validating license..."
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
    $form.Refresh()

    $key = $txtLicense.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($key)) {
        $lblStatus.Text      = "ERROR: Enter a License Key first."
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        return
    }

    $licenseUrl = "https://raw.githubusercontent.com/chakumon/entrypass-licenses/main/licenses.json"
    try {
        $response    = Invoke-WebRequest -Uri $licenseUrl -UseBasicParsing -TimeoutSec 10
        $licenseData = $response.Content | ConvertFrom-Json
        $entry = $licenseData | Where-Object { $_.licenseKey -eq $key }

        if (-not $entry) {
            $lblStatus.Text      = "License NOT FOUND. Check the key or contact Dajayana Trading."
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            return
        }
        if ($entry.active -ne $true) {
            $lblStatus.Text      = "License is INACTIVE. Contact Dajayana Trading at +60 16-883 8338"
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            return
        }
        $expiry = [datetime]::Parse($entry.expires)
        if ($expiry -lt (Get-Date)) {
            $lblStatus.Text      = "License EXPIRED on $($expiry.ToString('yyyy-MM-dd')). Contact Dajayana Trading."
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            return
        }
        $lblStatus.Text      = "License VALID. Client: $($entry.client). Expires: $($expiry.ToString('yyyy-MM-dd'))."
        $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
    } catch {
        $lblStatus.Text      = "Could not reach license server: $($_.Exception.Message)"
        $lblStatus.ForeColor = [System.Drawing.Color]::DarkOrange
    }
})

# ============================================================
# BUTTON -- Save Only
# ============================================================
$btnSaveOnly.Add_Click({
    if (-not (Validate-Fields)) { return }

    try {
        Save-Config | Out-Null
        $lblStatus.Text      = "config.json saved to: $scriptDir"
        $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
        [System.Windows.Forms.MessageBox]::Show(
            "Configuration saved successfully.`n`nFile: $configFile`n`nRun sync.ps1 manually or use 'Save && Install' to set up automatic scheduling.",
            "Saved",
            "OK",
            "Information"
        )
    } catch {
        $lblStatus.Text      = "ERROR saving config: $_"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
    }
})

# ============================================================
# BUTTON -- Save & Install
# ============================================================
$btnSaveInstall.Add_Click({
    if (-not (Validate-Fields)) { return }

    try {
        $cfg = Save-Config
        $lblStatus.Text      = "Config saved. Creating scheduled task..."
        $lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
        $form.Refresh()

        $taskResult = Install-ScheduledTask -LocationCode $cfg.location -Frequency $cmbFrequency.SelectedItem.ToString()

        if ($taskResult.Success) {
            $lblStatus.Text      = "Done! Config saved and task installed."
            $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
            [System.Windows.Forms.MessageBox]::Show(
                "Setup complete!`n`nConfig saved: $configFile`n`nScheduled task: EntryPass-MiHCM Sync - $($cfg.location)`nFrequency: $($cmbFrequency.SelectedItem)`n`nNote: The task runs as SYSTEM or current user. If it does not trigger, open Task Scheduler and verify credentials.",
                "Setup Complete",
                "OK",
                "Information"
            )
        } else {
            $lblStatus.Text      = "Config saved but task install failed. See details."
            $lblStatus.ForeColor = [System.Drawing.Color]::DarkOrange
            [System.Windows.Forms.MessageBox]::Show(
                "Config saved successfully, but the scheduled task could not be created:`n`n$($taskResult.Message)`n`nYou can create the task manually in Task Scheduler or run sync.ps1 via run.bat.",
                "Task Install Failed",
                "OK",
                "Warning"
            )
        }
    } catch {
        $lblStatus.Text      = "ERROR: $_"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
    }
})

# ============================================================
# SHOW FORM
# ============================================================
[void]$form.ShowDialog()
