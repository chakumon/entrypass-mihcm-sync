# sync.ps1 -- EntryPass to MiHCM Sync Engine
# Reads config from config.json in the same directory.
# Run manually via run.bat or automatically via Windows Task Scheduler.
#
# Requirements: PowerShell 5.1+, Windows 10/11
# No external dependencies required.

$ErrorActionPreference = "Continue"

# ============================================================
# SCRIPT DIRECTORY -- portable, works from Task Scheduler too
# ============================================================
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($scriptDir)) {
    $scriptDir = Split-Path -Parent $PSCommandPath
}
if ([string]::IsNullOrWhiteSpace($scriptDir)) {
    $scriptDir = Get-Location
}

# ============================================================
# LOGGING -- appends to sync_log.txt in the script directory
# ============================================================
$logFile = Join-Path $scriptDir "sync_log.txt"

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] $Message"
    Write-Host $entry
    Add-Content -Path $logFile -Value $entry -Encoding UTF8
}

# ============================================================
# RETRY HELPER -- wraps any scriptblock with exponential backoff
# ============================================================
function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 3
    )

    $retryableStatusCodes = @(429, 500, 502, 503, 504)
    $attempt = 0
    $lastError = $null

    while ($attempt -le $MaxRetries) {
        try {
            return (& $ScriptBlock)
        } catch {
            $lastError = $_
            $statusCode = $null

            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            if ($attempt -lt $MaxRetries -and $statusCode -in $retryableStatusCodes) {
                $waitSeconds = [Math]::Pow(2, $attempt + 1)   # 2s, 4s, 8s

                # Honour Retry-After on HTTP 429
                if ($statusCode -eq 429) {
                    try {
                        $retryAfterHeader = $_.Exception.Response.Headers["Retry-After"]
                        $parsed = 0
                        if ($retryAfterHeader -and [int]::TryParse($retryAfterHeader, [ref]$parsed) -and $parsed -gt 0) {
                            $waitSeconds = $parsed
                        }
                    } catch {}
                }

                Write-Log "RETRY: HTTP $statusCode on attempt $($attempt + 1)/$MaxRetries -- waiting ${waitSeconds}s before retry..."
                Start-Sleep -Seconds $waitSeconds
                $attempt++
            } else {
                throw $lastError
            }
        }
    }

    throw $lastError
}

# ============================================================
# LOAD CONFIG -- reads config.json from script directory
# ============================================================
$configFile = Join-Path $scriptDir "config.json"
if (-not (Test-Path $configFile)) {
    Write-Host "ERROR: config.json not found in $scriptDir"
    Write-Host "Please run setup.ps1 to configure the application."
    if ($Host.UI.RawUI.KeyAvailable -ne $null) { Read-Host "Press Enter to close..." }
    exit 1
}

try {
    $config = Get-Content $configFile -Raw -Encoding UTF8 | ConvertFrom-Json
} catch {
    Write-Host "ERROR: Failed to parse config.json -- $_"
    if ($Host.UI.RawUI.KeyAvailable -ne $null) { Read-Host "Press Enter to close..." }
    exit 1
}

# Validate required config fields
$requiredFields = @("licenseKey", "primaryKey", "secretKey", "apiEndpoint", "sourceFolder", "location")
foreach ($field in $requiredFields) {
    if ([string]::IsNullOrWhiteSpace($config.$field)) {
        Write-Host "ERROR: config.json is missing required field: $field"
        Write-Host "Please run setup.ps1 to reconfigure."
        if ($Host.UI.RawUI.KeyAvailable -ne $null) { Read-Host "Press Enter to close..." }
        exit 1
    }
}

# Assign config values to working variables
$MIHCM_PRIMARY_KEY = $config.primaryKey
$MIHCM_SECRET_KEY  = $config.secretKey
$MIHCM_BASE_URL    = $config.apiEndpoint.TrimEnd("/")
$BATCH_SIZE        = if ($config.batchSize -gt 0) { [int]$config.batchSize } else { 80 }
$LOCATION_CODE     = $config.location
$LOCATION_DESC     = if ($config.locationDesc) { $config.locationDesc } else { $config.location }
$CLIENT_NAME       = if ($config.clientName) { $config.clientName } else { "Unknown" }
$SOURCE_FOLDER     = $config.sourceFolder
$LICENSE_KEY       = $config.licenseKey

# Use sourceFolder from config; fall back to scriptDir if not set
if ([string]::IsNullOrWhiteSpace($SOURCE_FOLDER)) {
    $SOURCE_FOLDER = $scriptDir
}

# ============================================================
# LICENSE VALIDATION
# Checks online first, falls back to local cache if offline.
# ============================================================
$licenseUrl       = "https://raw.githubusercontent.com/chakumon/entrypass-licenses/main/licenses.json"
$licenseCacheFile = Join-Path $scriptDir "license_cache.json"

function Test-License {
    param([string]$Key)

    $licenseData = $null
    $source      = "online"

    # Attempt online check
    try {
        Write-Log "LICENSE: Checking license online..."
        $response = Invoke-WebRequest -Uri $licenseUrl -UseBasicParsing -TimeoutSec 10
        $licenseData = $response.Content | ConvertFrom-Json
        $source = "online"
    } catch {
        Write-Log "LICENSE: Online check failed ($($_.Exception.Message)) -- trying local cache..."
    }

    # Fall back to cache if online failed
    if (-not $licenseData) {
        if (Test-Path $licenseCacheFile) {
            try {
                $licenseData = Get-Content $licenseCacheFile -Raw -Encoding UTF8 | ConvertFrom-Json
                $source = "cache"
                Write-Log "LICENSE: Using cached license data."
            } catch {
                Write-Log "LICENSE: Cache file unreadable -- $_"
            }
        }
    }

    if (-not $licenseData) {
        Write-Log "LICENSE: No license data available (online or cache). Cannot validate."
        return $false
    }

    # Find entry matching the license key
    $entry = $licenseData | Where-Object { $_.licenseKey -eq $Key }
    if (-not $entry) {
        Write-Log "LICENSE: Key not found in license list ($source)."
        return $false
    }

    if ($entry.active -ne $true) {
        Write-Log "LICENSE: License is marked inactive ($source)."
        return $false
    }

    # Check expiry date
    try {
        $expiry = [datetime]::Parse($entry.expires)
        if ($expiry -lt (Get-Date)) {
            Write-Log "LICENSE: License expired on $($expiry.ToString('yyyy-MM-dd')) ($source)."
            return $false
        }
    } catch {
        Write-Log "LICENSE: Could not parse expiry date '$($entry.expires)' -- $_"
        return $false
    }

    Write-Log "LICENSE: Valid. Client: $($entry.client). Expires: $($entry.expires). Source: $source."

    # Save to local cache if we got a fresh online response
    if ($source -eq "online") {
        try {
            $licenseData | ConvertTo-Json -Depth 10 | Set-Content -Path $licenseCacheFile -Encoding UTF8
            Write-Log "LICENSE: Cache updated."
        } catch {
            Write-Log "LICENSE: Warning -- could not write cache: $_"
        }
    }

    return $true
}

# ============================================================
# LOG HEADER
# ============================================================
Add-Content -Path $logFile -Value "" -Encoding UTF8
Add-Content -Path $logFile -Value "========================================" -Encoding UTF8
Write-Log "EntryPass-MiHCM Sync"
Write-Log "Client   : $CLIENT_NAME"
Write-Log "Location : $LOCATION_CODE ($LOCATION_DESC)"
Write-Log "Source   : $SOURCE_FOLDER"
Write-Log "Endpoint : $MIHCM_BASE_URL"

# ============================================================
# VALIDATE LICENSE
# ============================================================
$licenseValid = Test-License -Key $LICENSE_KEY
if (-not $licenseValid) {
    Write-Log "LICENSE: Invalid or expired. Exiting."
    Write-Host ""
    Write-Host "License invalid or expired. Contact Dajayana Trading at +60 16-883 8338"
    if ($Host.UI.RawUI.KeyAvailable -ne $null) { Read-Host "Press Enter to close..." }
    exit 1
}

# ============================================================
# GET MIHCM ACCESS TOKEN
# ============================================================
function Get-MiHCMToken {
    $tokenUrl = "$MIHCM_BASE_URL/oauth2/token?grantType=client_credentials&clientId=$MIHCM_PRIMARY_KEY&clientSecret=$MIHCM_SECRET_KEY"
    Write-Log "API >> GET $MIHCM_BASE_URL/oauth2/token"
    try {
        $raw = Invoke-WithRetry -ScriptBlock {
            Invoke-WebRequest -Uri $tokenUrl -Method GET -Headers @{
                "Ocp-Apim-Subscription-Key" = $MIHCM_PRIMARY_KEY
            } -UseBasicParsing
        }
        Write-Log "API << HTTP $($raw.StatusCode) $($raw.StatusDescription)"
        $response = $raw.Content | ConvertFrom-Json
        if ($response.accessToken) {
            Write-Log "API << Token obtained. Expires: $($response.expiresOn)"
            return $response.accessToken
        } else {
            Write-Log "API << ERROR: No accessToken in response body"
            Write-Log "API << Body: $($raw.Content)"
            return $null
        }
    } catch {
        Write-Log "API << ERROR: $($_.Exception.Message)"
        if ($_.Exception.Response) {
            $statusCode = [int]$_.Exception.Response.StatusCode
            Write-Log "API << HTTP $statusCode"
            try {
                $reader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
                $errorBody = $reader.ReadToEnd()
                $reader.Close()
                Write-Log "API << Body: $errorBody"
            } catch {}
        }
        return $null
    }
}

# ============================================================
# UPLOAD RECORDS TO MIHCM -- batched, max 100 per request
# ============================================================
function Upload-ToMiHCM {
    param(
        [string]$AccessToken,
        [array]$Records
    )

    $headers = @{
        "Ocp-Apim-Subscription-Key" = $MIHCM_PRIMARY_KEY
        "Authorization"             = "Bearer $AccessToken"
        "Content-Type"              = "application/json"
    }

    $uploadUrl  = "$MIHCM_BASE_URL/ontime/clockfileuploads"
    $totalSaved   = 0
    $totalSkipped = 0
    $totalFailed  = 0

    # Split records into batches
    $batches = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $Records.Count; $i += $BATCH_SIZE) {
        $end = [Math]::Min($i + $BATCH_SIZE, $Records.Count)
        $batches.Add($Records[$i..($end - 1)])
    }

    Write-Log "Uploading $($Records.Count) records in $($batches.Count) batch(es) (size: $BATCH_SIZE)..."

    $batchNum = 1
    foreach ($batch in $batches) {
        Write-Log "API >> POST $uploadUrl  [batch $batchNum, $($batch.Count) records]"
        $body = $batch | ConvertTo-Json -Depth 5

        # Log one sample record per batch for verification
        $sampleRecord = ($batch | Select-Object -First 1) | ConvertTo-Json -Compress -Depth 5
        Write-Log "API >> Sample: $sampleRecord"

        $batchUploaded  = $false
        $reAuthAttempted = $false

        :batchRetry while (-not $batchUploaded) {
            try {
                $raw = Invoke-WithRetry -ScriptBlock {
                    Invoke-WebRequest -Uri $uploadUrl -Method POST -Headers $headers -Body $body -UseBasicParsing
                }
                Write-Log "API << HTTP $($raw.StatusCode) $($raw.StatusDescription)"
                Write-Log "API << Body: $($raw.Content)"

                $response = $raw.Content | ConvertFrom-Json
                if ($response.statusDetail) {
                    foreach ($detail in $response.statusDetail) {
                        if ($detail.success -eq $true) {
                            $totalSaved++
                        } elseif ($detail.message -match "already exists") {
                            $totalSkipped++
                        } else {
                            $totalFailed++
                            Write-Log "    WARN: $($detail.textCardNumber) @ $($detail.date) -- $($detail.message)"
                        }
                    }
                    $bSaved   = ($response.statusDetail | Where-Object { $_.success -eq $true }).Count
                    $bSkipped = ($response.statusDetail | Where-Object { $_.message -match "already exists" }).Count
                    $bFailed  = ($response.statusDetail | Where-Object { $_.success -ne $true -and $_.message -notmatch "already exists" }).Count
                    Write-Log "  Batch $batchNum -- Saved: $bSaved | Skipped: $bSkipped | Failed: $bFailed"
                } else {
                    Write-Log "  Batch $batchNum response (no statusDetail): $($raw.Content)"
                }
                $batchUploaded = $true

            } catch {
                Write-Log "API << ERROR: $($_.Exception.Message)"
                $batchStatusCode = $null
                if ($_.Exception.Response) {
                    $batchStatusCode = [int]$_.Exception.Response.StatusCode
                    Write-Log "API << HTTP $batchStatusCode"
                    try {
                        $reader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
                        $errorBody = $reader.ReadToEnd()
                        $reader.Close()
                        Write-Log "API << Body: $errorBody"
                    } catch {}
                }

                # 401 Unauthorized -- re-fetch token once, then retry this batch
                if ($batchStatusCode -eq 401 -and -not $reAuthAttempted) {
                    Write-Log "AUTH: HTTP 401 on batch $batchNum -- re-fetching token..."
                    $reAuthAttempted = $true
                    $newToken = Get-MiHCMToken
                    if ($newToken) {
                        $headers["Authorization"] = "Bearer $newToken"
                        Write-Log "AUTH: Token refreshed -- retrying batch $batchNum..."
                        continue batchRetry
                    } else {
                        Write-Log "AUTH: Token refresh failed -- batch $batchNum marked as failed."
                    }
                }

                $totalFailed  += $batch.Count
                $batchUploaded = $true   # stop retrying this batch
            }
        }
        $batchNum++
    }

    Write-Log "Upload complete -- Saved: $totalSaved | Skipped (duplicate): $totalSkipped | Failed: $totalFailed"
    return @{ Saved = $totalSaved; Skipped = $totalSkipped; Failed = $totalFailed }
}

# ============================================================
# CONVERT A SINGLE EntryPass DATA FILE
# ============================================================
function Convert-EntryPassFile {
    param([string]$InputFile)

    Write-Log "Processing: $InputFile"

    # Auto-detect delimiter (semicolon or comma)
    $firstLine = Get-Content $InputFile -Encoding UTF8 | Select-Object -First 1
    $delimiter = if ($firstLine -match ';') { ';' } else { ',' }
    Write-Log "  Detected delimiter: $delimiter"

    $lines   = Get-Content $InputFile -Encoding UTF8
    $records = @()

    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        $parts = $line -split $delimiter
        if ($parts.Count -lt 3) { continue }

        $dateStr    = $parts[0].Trim()
        $timeStr    = $parts[1].Trim()
        $cardNumber = $parts[2].Trim() -replace '\.0*$', ''

        if ([string]::IsNullOrWhiteSpace($cardNumber)) { continue }

        # Convert date: YYYY/MM/DD -> YYYY-MM-DD
        $dateFormatted = $dateStr -replace '(\d{4})/(\d{2})/(\d{2})', '$1-$2-$3'
        $dateOnly = $dateFormatted

        # Build datetime strings for MiHCM format: "YYYY-MM-DD HH:mm:ss.000"
        $dateFull = "$dateOnly 00:00:00.000"

        # Handle both HH:MM and HH:MM:SS time formats from EntryPass
        if ($timeStr -match '^\d{2}:\d{2}:\d{2}$') {
            $timeFull = "$dateOnly $($timeStr).000"
        } else {
            $timeFull = "$dateOnly $($timeStr):00.000"
        }

        $records += @{
            "Date"           = $dateFull
            "Time"           = $timeFull
            "CardNumber"     = 0
            "Node"           = 0
            "TextCardNumber" = $cardNumber
            "Clock"          = 0
            "TrType"         = 0
            "Location"       = $LOCATION_CODE
        }
    }

    if ($records.Count -eq 0) {
        Write-Log "  No valid records found -- skipping."
        return $null
    }

    Write-Log "  Parsed $($records.Count) records."

    # Save a local backup copy alongside the source file
    $baseName   = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
    $dateTag    = $baseName -replace 'DATA', ''
    $backupFile = Join-Path (Split-Path $InputFile -Parent) "attendance_mihcm_$dateTag.txt"
    $backupLines = $records | ForEach-Object {
        "$($_['TextCardNumber'])`t$($_['Date'].Substring(0,10))`t$($_['Time'].Substring(11,5))"
    }
    [System.IO.File]::WriteAllLines($backupFile, $backupLines)
    Write-Log "  Backup saved: $backupFile"

    return $records
}

# ============================================================
# MAIN -- find files, get token, convert, upload, summarise
# ============================================================
$runStart = Get-Date

# Find all DATA*.txt files in the configured source folder
$files = Get-ChildItem -Path $SOURCE_FOLDER -Filter "DATA*.txt" -ErrorAction SilentlyContinue

if ($files.Count -eq 0) {
    Write-Log "No DATA*.txt files found in $SOURCE_FOLDER. Nothing to do."
    if ($Host.UI.RawUI -and $Host.Name -ne "Default Host") { Read-Host "Press Enter to close..." }
    exit 0
}

Write-Log "Found $($files.Count) file(s) to process."

# Get MiHCM access token (valid ~1 hour -- sufficient for one run)
$token = Get-MiHCMToken
if (-not $token) {
    Write-Log "Cannot proceed without a valid token. Check your API keys."
    if ($Host.UI.RawUI -and $Host.Name -ne "Default Host") { Read-Host "Press Enter to close..." }
    exit 1
}

# Convert each file and collect all records
$allRecords = @()
foreach ($file in $files) {
    $records = Convert-EntryPassFile -InputFile $file.FullName
    if ($records) {
        $allRecords += $records
    }
    Write-Log ""
}

if ($allRecords.Count -eq 0) {
    Write-Log "No records to upload after parsing all files."
    if ($Host.UI.RawUI -and $Host.Name -ne "Default Host") { Read-Host "Press Enter to close..." }
    exit 0
}

Write-Log "Total records to upload: $($allRecords.Count)"
$result = Upload-ToMiHCM -AccessToken $token -Records $allRecords

# Delete original DATA files only when upload had zero failures
if ($result.Failed -eq 0) {
    foreach ($file in $files) {
        try {
            Remove-Item $file.FullName -Force
            Write-Log "Deleted source file: $($file.Name)"
        } catch {
            Write-Log "Warning: Could not delete $($file.Name) -- $_"
        }
    }
} else {
    Write-Log "WARNING: $($result.Failed) record(s) failed -- source files NOT deleted. Check manually."
}

$runEnd      = Get-Date
$runDuration = ($runEnd - $runStart).TotalSeconds

# ============================================================
# RUN SUMMARY
# ============================================================
$totalRecordsParsed = $allRecords.Count
$totalSavedFinal    = if ($result) { $result.Saved }   else { 0 }
$totalSkippedFinal  = if ($result) { $result.Skipped } else { 0 }
$totalFailedFinal   = if ($result) { $result.Failed }  else { 0 }

if ($totalFailedFinal -eq 0 -and $totalSavedFinal -gt 0) {
    $overallResult = "SUCCESS"
} elseif ($totalFailedFinal -gt 0 -and $totalSavedFinal -gt 0) {
    $overallResult = "PARTIAL FAILURE"
} elseif ($totalSavedFinal -eq 0 -and $totalSkippedFinal -gt 0) {
    $overallResult = "ALL DUPLICATES (no new records)"
} else {
    $overallResult = "FAILURE"
}

Add-Content -Path $logFile -Value "" -Encoding UTF8
Write-Log "========================================"
Write-Log "RUN SUMMARY"
Write-Log "  Client           : $CLIENT_NAME"
Write-Log "  Location         : $LOCATION_CODE"
Write-Log "  Files processed  : $($files.Count)"
Write-Log "  Records parsed   : $totalRecordsParsed"
Write-Log "  Uploaded (saved) : $totalSavedFinal"
Write-Log "  Skipped (dupes)  : $totalSkippedFinal"
Write-Log "  Failed           : $totalFailedFinal"
Write-Log "  Start time       : $($runStart.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Log "  End time         : $($runEnd.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Log "  Duration         : $([Math]::Round($runDuration, 1))s"
Write-Log "  Overall result   : $overallResult"
Write-Log "========================================"
Write-Log "Done."

Write-Host ""
# Only prompt in interactive sessions (not when run by Task Scheduler)
if ($Host.UI.RawUI -and $Host.Name -ne "Default Host") {
    Read-Host "Press Enter to close..."
}
