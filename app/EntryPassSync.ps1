# EntryPassSync.ps1 -- EntryPass to MiHCM Sync -- All-in-One GUI Application
# Company  : Dajayana Trading (www.dajayana.com)
# Version  : 1.0.0
# Contact  : +60 16-883 8338
# Requires : PowerShell 5.1+, Windows 10/11

$script:appVersion = "1.0.0"
$script:updateUrl  = "https://raw.githubusercontent.com/chakumon/entrypass-mihcm-sync/main/version.json"
$script:scriptUrl  = "https://raw.githubusercontent.com/chakumon/entrypass-mihcm-sync/main/app/EntryPassSync.ps1"

$ErrorActionPreference = "Continue"

# Hide the PowerShell console window
Add-Type -Name Win32 -Namespace Native -MemberDefinition @"
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
"@
$consoleWnd = [Native.Win32]::GetConsoleWindow()
if ($consoleWnd -ne [IntPtr]::Zero) { [Native.Win32]::ShowWindow($consoleWnd, 0) | Out-Null }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# ============================================================
# SCRIPT DIRECTORY
# ============================================================
$script:appDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($script:appDir)) { $script:appDir = Split-Path -Parent $PSCommandPath }
if ([string]::IsNullOrWhiteSpace($script:appDir)) { $script:appDir = (Get-Location).Path }

$script:configFile  = Join-Path $script:appDir "config.json"
$script:logFile     = Join-Path $script:appDir "sync_log.txt"
$script:cacheFile   = Join-Path $script:appDir "license_cache.json"
$script:licenseUrl  = "https://raw.githubusercontent.com/chakumon/entrypass-mihcm-licenses/main/licenses.json"
$script:versionFile = Join-Path $script:appDir "version_cache.json"

# ============================================================
# SCRIPT-LEVEL STATE
# ============================================================
$script:bgWorker       = $null
$script:reallyExit     = $false
$script:lastStats      = @{ Saved = 0; Skipped = 0; Failed = 0; Result = ""; Time = "" }
$script:activePanelName = ""
$script:navItems       = @{}

# Color palette
$clrSidebar    = [System.Drawing.Color]::FromArgb(26,35,50)
$clrSidebarAct = [System.Drawing.Color]::FromArgb(30,72,120)
$clrSidebarHov = [System.Drawing.Color]::FromArgb(42,58,80)
$clrPanelBg    = [System.Drawing.Color]::FromArgb(240,242,245)
$clrDarkBox    = [System.Drawing.Color]::FromArgb(15,22,34)
$clrGreen      = [System.Drawing.Color]::FromArgb(56,180,100)
$clrGrey       = [System.Drawing.Color]::FromArgb(120,130,148)
$clrBlue       = [System.Drawing.Color]::FromArgb(0,102,204)
$clrTextDim    = [System.Drawing.Color]::FromArgb(90,100,115)
$clrTextDark   = [System.Drawing.Color]::FromArgb(40,50,70)
$clrOrange     = [System.Drawing.Color]::FromArgb(210,140,40)

# ============================================================
# CONFIG HELPERS
# ============================================================
function Load-AppConfig {
    if (Test-Path $script:configFile) {
        try { return Get-Content $script:configFile -Raw -Encoding UTF8 | ConvertFrom-Json } catch {}
    }
    return $null
}

function Save-AppConfig {
    param($Cfg)
    $json = $Cfg | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($script:configFile, $json, [System.Text.Encoding]::UTF8)
}

function Is-Configured {
    $cfg = Load-AppConfig
    if (-not $cfg) { return $false }
    $req = @("licenseKey","primaryKey","secretKey","location")
    foreach ($f in $req) {
        if ([string]::IsNullOrWhiteSpace($cfg.$f)) { return $false }
    }
    if ($cfg.dataSource -eq "database") {
        if ([string]::IsNullOrWhiteSpace($cfg.databasePath)) { return $false }
    } else {
        if ([string]::IsNullOrWhiteSpace($cfg.sourceFolder)) { return $false }
    }
    return $true
}

# ============================================================
# SYNC ENGINE FUNCTIONS
# ============================================================
function Write-SyncLog {
    param([string]$Message)
    $ts    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$ts] $Message"
    try { Add-Content -Path $script:logFile -Value $entry -Encoding UTF8 } catch {}
    if ($script:txtLiveLog -ne $null) {
        try {
            $script:txtLiveLog.AppendText("$entry`r`n")
            $script:txtLiveLog.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        } catch {}
    }
}

function Trim-LogFile {
    # Remove log entries older than 90 days
    $maxAge = 90
    if (-not (Test-Path $script:logFile)) { return }
    try {
        $cutoff = (Get-Date).AddDays(-$maxAge).ToString("yyyy-MM-dd")
        $lines  = [System.IO.File]::ReadAllLines($script:logFile, [System.Text.Encoding]::UTF8)
        $kept   = [System.Collections.Generic.List[string]]::new()
        foreach ($line in $lines) {
            # Log lines start with [YYYY-MM-DD HH:MM:SS]
            if ($line -match '^\[(\d{4}-\d{2}-\d{2})') {
                if ($Matches[1] -ge $cutoff) { $kept.Add($line) }
            } elseif ($line -match '^={5,}') {
                # Separator lines -- keep if following entries are recent
                $kept.Add($line)
            } else {
                # Non-dated lines (blank, continuation) -- keep
                $kept.Add($line)
            }
        }
        $removed = $lines.Count - $kept.Count
        if ($removed -gt 0) {
            [System.IO.File]::WriteAllLines($script:logFile, $kept.ToArray(), [System.Text.Encoding]::UTF8)
            Write-SyncLog "Log cleanup: removed $removed entries older than $maxAge days"
        }
    } catch {
        # Silently ignore cleanup errors -- not critical
    }
}

function Invoke-WithRetry {
    param([scriptblock]$ScriptBlock, [int]$MaxRetries = 3)
    $retryable = @(429,500,502,503,504)
    $attempt   = 0
    $lastErr   = $null
    while ($attempt -le $MaxRetries) {
        try { return (& $ScriptBlock) } catch {
            $lastErr = $_
            $sc      = $null
            if ($_.Exception.Response) { $sc = [int]$_.Exception.Response.StatusCode }
            if ($attempt -lt $MaxRetries -and $sc -in $retryable) {
                $wait = [Math]::Pow(2, $attempt + 1)
                if ($sc -eq 429) {
                    try {
                        $ra  = $_.Exception.Response.Headers["Retry-After"]
                        $p   = 0
                        if ($ra -and [int]::TryParse($ra,[ref]$p) -and $p -gt 0) { $wait = $p }
                    } catch {}
                }
                Write-SyncLog "RETRY: HTTP $sc attempt $($attempt+1)/$MaxRetries -- waiting ${wait}s..."
                Start-Sleep -Seconds $wait
                $attempt++
            } else { throw $lastErr }
        }
    }
    throw $lastErr
}

function Test-LicenseKey {
    param([string]$Key)
    $data   = $null
    $source = "online"
    try {
        Write-SyncLog "LICENSE: Checking online..."
        $r    = Invoke-WebRequest -Uri $script:licenseUrl -UseBasicParsing -TimeoutSec 10
        $data = $r.Content | ConvertFrom-Json
    } catch {
        Write-SyncLog "LICENSE: Online failed ($($_.Exception.Message)) -- trying cache..."
    }
    if (-not $data) {
        if (Test-Path $script:cacheFile) {
            try { $data = Get-Content $script:cacheFile -Raw -Encoding UTF8 | ConvertFrom-Json; $source = "cache" } catch {}
        }
    }
    if (-not $data) { Write-SyncLog "LICENSE: No data available."; return $false }
    # License JSON is a dictionary keyed by license key
    $entry = $data.$Key
    if (-not $entry) { $entry = $data.PSObject.Properties[$Key].Value 2>$null }
    if (-not $entry) { Write-SyncLog "LICENSE: Key not found ($source)."; return $false }
    if ($entry.active -ne $true) { Write-SyncLog "LICENSE: Key is inactive ($source)."; return $false }
    try {
        $exp = [datetime]::Parse($entry.expires)
        if ($exp -lt (Get-Date)) { Write-SyncLog "LICENSE: Expired $($exp.ToString('yyyy-MM-dd')) ($source)."; return $false }
    } catch { Write-SyncLog "LICENSE: Cannot parse expiry -- $_"; return $false }
    Write-SyncLog "LICENSE: Valid. Client=$($entry.client) Expires=$($entry.expires) Source=$source"
    if ($source -eq "online") {
        try { $data | ConvertTo-Json -Depth 10 | Set-Content -Path $script:cacheFile -Encoding UTF8 } catch {}
    }
    return $true
}

function Get-MiHCMToken {
    param([string]$BaseUrl,[string]$PrimaryKey,[string]$SecretKey)
    $url = "$BaseUrl/oauth2/token?grantType=client_credentials&clientId=$PrimaryKey&clientSecret=$SecretKey"
    Write-SyncLog "API >> GET $BaseUrl/oauth2/token"
    try {
        $raw  = Invoke-WithRetry -ScriptBlock { Invoke-WebRequest -Uri $url -Method GET -Headers @{"Ocp-Apim-Subscription-Key"=$PrimaryKey} -UseBasicParsing }
        Write-SyncLog "API << HTTP $($raw.StatusCode)"
        $resp = $raw.Content | ConvertFrom-Json
        if ($resp.accessToken) { Write-SyncLog "API << Token obtained."; return $resp.accessToken }
        Write-SyncLog "API << ERROR: no accessToken. Body: $($raw.Content)"
        return $null
    } catch {
        Write-SyncLog "API << ERROR: $($_.Exception.Message)"
        if ($_.Exception.Response) {
            try { $r=[System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream()); Write-SyncLog "API << Body: $($r.ReadToEnd())"; $r.Close() } catch {}
        }
        return $null
    }
}

function Upload-Records {
    param([string]$BaseUrl,[string]$PrimaryKey,[string]$SecretKey,[string]$Token,[array]$Records,[int]$BatchSize)
    $hdrs = @{
        "Ocp-Apim-Subscription-Key" = $PrimaryKey
        "Authorization"             = "Bearer $Token"
        "Content-Type"              = "application/json"
    }
    $url         = "$BaseUrl/ontime/clockfileuploads"
    $totalSaved  = 0; $totalSkip = 0; $totalFail = 0

    $batches = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $Records.Count; $i += $BatchSize) {
        $end = [Math]::Min($i + $BatchSize, $Records.Count)
        $batches.Add($Records[$i..($end-1)])
    }
    Write-SyncLog "Uploading $($Records.Count) records in $($batches.Count) batch(es) (size=$BatchSize)..."

    $bNum = 1
    foreach ($batch in $batches) {
        Write-SyncLog "API >> POST $url  [batch $bNum, $($batch.Count) records]"
        $bodyJson = $batch | ConvertTo-Json -Depth 5
        $sample   = ($batch | Select-Object -First 1) | ConvertTo-Json -Compress -Depth 5
        Write-SyncLog "API >> Sample: $sample"
        $done    = $false
        $reAuthed = $false
        :retry while (-not $done) {
            try {
                $raw  = Invoke-WithRetry -ScriptBlock { Invoke-WebRequest -Uri $url -Method POST -Headers $hdrs -Body $bodyJson -UseBasicParsing }
                Write-SyncLog "API << HTTP $($raw.StatusCode)"
                $resp = $raw.Content | ConvertFrom-Json
                if ($resp.statusDetail) {
                    foreach ($d in $resp.statusDetail) {
                        if ($d.success -eq $true) { $totalSaved++ }
                        elseif ($d.message -match "already exists") { $totalSkip++ }
                        else { $totalFail++; Write-SyncLog "    WARN: $($d.textCardNumber) @ $($d.date) -- $($d.message)" }
                    }
                    $bS = ($resp.statusDetail | Where-Object {$_.success -eq $true}).Count
                    $bK = ($resp.statusDetail | Where-Object {$_.message -match "already exists"}).Count
                    $bF = ($resp.statusDetail | Where-Object {$_.success -ne $true -and $_.message -notmatch "already exists"}).Count
                    Write-SyncLog "  Batch $bNum -- Saved:$bS Skipped:$bK Failed:$bF"
                } else {
                    Write-SyncLog "  Batch $bNum response: $($raw.Content)"
                }
                $done = $true
            } catch {
                Write-SyncLog "API << ERROR: $($_.Exception.Message)"
                $bsc = $null
                if ($_.Exception.Response) {
                    $bsc = [int]$_.Exception.Response.StatusCode
                    try { $r=[System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream()); Write-SyncLog "API << Body: $($r.ReadToEnd())"; $r.Close() } catch {}
                }
                if ($bsc -eq 401 -and -not $reAuthed) {
                    $reAuthed = $true
                    Write-SyncLog "AUTH: HTTP 401 -- re-fetching token..."
                    $newTok = Get-MiHCMToken -BaseUrl $BaseUrl -PrimaryKey $PrimaryKey -SecretKey $SecretKey
                    if ($newTok) { $hdrs["Authorization"] = "Bearer $newTok"; Write-SyncLog "AUTH: Token refreshed."; continue retry }
                    Write-SyncLog "AUTH: Token refresh failed."
                }
                $totalFail += $batch.Count
                $done       = $true
            }
        }
        $bNum++
    }
    Write-SyncLog "Upload complete -- Saved:$totalSaved Skipped:$totalSkip Failed:$totalFail"
    return @{ Saved=$totalSaved; Skipped=$totalSkip; Failed=$totalFail }
}

function Convert-EntryPassFile {
    param([string]$InputFile,[string]$LocationCode)
    Write-SyncLog "Processing: $InputFile"
    $firstLine = Get-Content $InputFile -Encoding UTF8 | Select-Object -First 1
    $delim     = if ($firstLine -match ';') { ';' } else { ',' }
    Write-SyncLog "  Delimiter: $delim"
    $lines   = Get-Content $InputFile -Encoding UTF8
    $records = @()
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $parts = $line -split $delim
        if ($parts.Count -lt 3) { continue }
        $dateStr = $parts[0].Trim()
        $timeStr = $parts[1].Trim()
        $card    = $parts[2].Trim() -replace '\.0*$',''
        if ([string]::IsNullOrWhiteSpace($card)) { continue }
        $dateFmt = $dateStr -replace '(\d{4})/(\d{2})/(\d{2})','$1-$2-$3'
        $dateFull = "$dateFmt 00:00:00.000"
        if ($timeStr -match '^\d{2}:\d{2}:\d{2}$') {
            $timeFull = "$dateFmt $timeStr.000"
        } else {
            $timeFull = "$dateFmt $($timeStr):00.000"
        }
        $records += @{
            "Date"           = $dateFull
            "Time"           = $timeFull
            "CardNumber"     = 0
            "Node"           = 0
            "TextCardNumber" = $card
            "Clock"          = 0
            "TrType"         = 0
            "Location"       = $LocationCode
        }
    }
    if ($records.Count -eq 0) { Write-SyncLog "  No valid records."; return $null }
    Write-SyncLog "  Parsed $($records.Count) records."
    $base    = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
    $tag     = $base -replace 'DATA',''
    $backup  = Join-Path (Split-Path $InputFile -Parent) "attendance_mihcm_$tag.txt"
    $bLines  = $records | ForEach-Object { "$($_['TextCardNumber'])`t$($_['Date'].Substring(0,10))`t$($_['Time'].Substring(11,5))" }
    [System.IO.File]::WriteAllLines($backup, $bLines)
    Write-SyncLog "  Backup: $backup"
    return $records
}

function Load-FirebirdAssembly {
    param([string]$DbPath, [string]$ClientLibrary)
    # Try explicit clientLibrary path first, then search common locations
    $searchPaths = @()
    if (-not [string]::IsNullOrWhiteSpace($ClientLibrary) -and (Test-Path $ClientLibrary)) {
        $searchPaths += $ClientLibrary
    }
    # App folder first (bundled embedded engine + .NET provider)
    $searchPaths += @(
        (Join-Path $script:appDir "FirebirdSql.Data.FirebirdClient.dll")
    )
    foreach ($path in $searchPaths) {
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        try {
            [void][System.Reflection.Assembly]::LoadFrom($path)
            Write-SyncLog "Firebird: Loaded .NET provider from $path"
            return "dotnet"
        } catch { continue }
    }
    # Try GAC
    try {
        [void][System.Reflection.Assembly]::Load("FirebirdSql.Data.FirebirdClient")
        Write-SyncLog "Firebird: Loaded .NET provider from GAC"
        return "dotnet"
    } catch {}
    # Try ODBC as fallback
    try {
        $testConn = New-Object System.Data.Odbc.OdbcConnection
        $testConn.Dispose()
        Write-SyncLog "Firebird: Will use ODBC driver (FirebirdSql.Data.FirebirdClient.dll not found)"
        return "odbc"
    } catch {}
    Write-SyncLog "Firebird: No .NET provider or ODBC available."
    return $null
}

function Read-FirebirdDatabase {
    param(
        [string]$DbPath,
        [string]$FbUser,
        [string]$FbPassword,
        [string]$ClientLibrary,
        [string]$LocationCode,
        [int]$SyncDays = 1
    )

    $allRecords    = @()
    $totalRaw      = 0
    $totalValid    = 0
    $totalSkipped  = 0

    $dbFile = Join-Path $DbPath "event\TRANS.FDB"
    if (-not (Test-Path $dbFile)) {
        Write-SyncLog "ERROR: Database not found at $dbFile"
        return $null
    }
    Write-SyncLog "Data source: Database Direct (TRANS.FDB)"
    Write-SyncLog "Database   : $dbFile"

    $providerType = Load-FirebirdAssembly -DbPath $DbPath -ClientLibrary $ClientLibrary
    if (-not $providerType) {
        Write-SyncLog "ERROR: Cannot connect to Firebird -- no driver available."
        Write-SyncLog "  Install FirebirdSql.Data.FirebirdClient NuGet package or place the DLL next to this script."
        return $null
    }

    for ($dayOffset = 0; $dayOffset -lt $SyncDays; $dayOffset++) {
        $targetDate  = (Get-Date).AddDays(-$dayOffset)
        $dateStrTbl  = $targetDate.ToString("yyyyMMdd")
        $tableName   = "STT$dateStrTbl"
        Write-SyncLog "Querying table $tableName..."

        $conn = $null
        $reader = $null
        try {
            if ($providerType -eq "dotnet") {
                # ServerType=0 = server mode (connects to running Firebird service)
                $connStr = "Database=$dbFile;User=$FbUser;Password=$FbPassword;ServerType=0;DataSource=localhost;Charset=UTF8"
                $conn    = New-Object FirebirdSql.Data.FirebirdClient.FbConnection($connStr)
                $conn.Open()
                $cmd = $conn.CreateCommand()
                $cmd.CommandText = "SELECT DATA FROM $tableName WHERE DATA LIKE '0;%'"
                $reader = $cmd.ExecuteReader()
            } else {
                # ODBC fallback
                $connStr = "Driver={Firebird/InterBase(r) driver};Database=$dbFile;UID=$FbUser;PWD=$FbPassword;CHARSET=UTF8"
                $conn    = New-Object System.Data.Odbc.OdbcConnection($connStr)
                $conn.Open()
                $cmd = $conn.CreateCommand()
                $cmd.CommandText = "SELECT DATA FROM $tableName WHERE DATA LIKE '0;%'"
                $reader = $cmd.ExecuteReader()
            }

            $dayRaw   = 0
            $dayValid = 0
            $daySkip  = 0

            while ($reader.Read()) {
                $dataStr = $reader.GetString(0)
                $dayRaw++
                $parts = $dataStr -split ';'
                if ($parts.Count -lt 9) { $daySkip++; continue }
                if ($parts[0].Trim() -ne '0')  { $daySkip++; continue }
                if ($parts[3].Trim() -ne 'Ca') { $daySkip++; continue }
                $cardNo = $parts[8].Trim()
                if ([string]::IsNullOrWhiteSpace($cardNo)) { $daySkip++; continue }
                # Remove leading zeros from card number (or keep as-is -- MiHCM uses TextCardNumber)
                $rawDate = $parts[1].Trim()  # YYYY/MM/DD
                $rawTime = $parts[2].Trim()  # HH:MM:SS
                # Convert date from YYYY/MM/DD to YYYY-MM-DD
                $dateFmt = $rawDate -replace '(\d{4})/(\d{2})/(\d{2})','$1-$2-$3'
                $dateFull = "$dateFmt 00:00:00.000"
                if ($rawTime -match '^\d{2}:\d{2}:\d{2}$') {
                    $timeFull = "$dateFmt $rawTime.000"
                } else {
                    $timeFull = "$dateFmt $($rawTime):00.000"
                }
                $allRecords += @{
                    "Date"           = $dateFull
                    "Time"           = $timeFull
                    "CardNumber"     = 0
                    "Node"           = 0
                    "TextCardNumber" = $cardNo
                    "Clock"          = 0
                    "TrType"         = 0
                    "Location"       = $LocationCode
                }
                $dayValid++
            }
            $daySkip = $dayRaw - $dayValid
            $totalRaw    += $dayRaw
            $totalValid  += $dayValid
            $totalSkipped += $daySkip
            Write-SyncLog "Found $dayRaw raw events, $dayValid valid card swipes"
            Write-SyncLog "Skipping $daySkip non-attendance events (door opens, etc.)"
        } catch {
            Write-SyncLog "ERROR querying ${tableName}: $($_.Exception.Message)"
            # Table may not exist for this date -- not a fatal error
        } finally {
            if ($reader) { try { $reader.Close() } catch {} }
            if ($conn)   { try { $conn.Close(); $conn.Dispose() } catch {} }
        }
    }

    Write-SyncLog "Database read complete -- Total raw:$totalRaw Valid:$totalValid Skipped:$totalSkipped"
    return $allRecords
}

function Test-FirebirdConnection {
    param([string]$DbPath, [string]$FbUser, [string]$FbPassword, [string]$ClientLibrary)
    $dbFile = Join-Path $DbPath "event\TRANS.FDB"
    if (-not (Test-Path $dbFile)) {
        return @{ Success=$false; Message="Database not found: $dbFile" }
    }
    $providerType = Load-FirebirdAssembly -DbPath $DbPath -ClientLibrary $ClientLibrary
    if (-not $providerType) {
        return @{ Success=$false; Message="No Firebird driver found. Place FirebirdSql.Data.FirebirdClient.dll next to this script, or install the ODBC driver." }
    }
    $dateStrTbl = (Get-Date).ToString("yyyyMMdd")
    $tableName  = "STT$dateStrTbl"
    $conn   = $null
    $reader = $null
    try {
        if ($providerType -eq "dotnet") {
            $connStr = "Database=$dbFile;User=$FbUser;Password=$FbPassword;ServerType=0;DataSource=localhost;Charset=UTF8"
            $conn    = New-Object FirebirdSql.Data.FirebirdClient.FbConnection($connStr)
        } else {
            $connStr = "Driver={Firebird/InterBase(r) driver};Database=$dbFile;UID=$FbUser;PWD=$FbPassword;CHARSET=UTF8"
            $conn    = New-Object System.Data.Odbc.OdbcConnection($connStr)
        }
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT COUNT(*) FROM $tableName"
        $count = $cmd.ExecuteScalar()
        return @{ Success=$true; Message="Connected - $tableName found, $count records" }
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match "Table unknown" -or $msg -match "object.*not found" -or $msg -match "335544580") {
            return @{ Success=$true; Message="Connected to TRANS.FDB (no table $tableName yet -- no data for today)" }
        }
        return @{ Success=$false; Message="Connection failed: $msg" }
    } finally {
        if ($reader) { try { $reader.Close() } catch {} }
        if ($conn)   { try { $conn.Close(); $conn.Dispose() } catch {} }
    }
}

function Run-FullSync {
    param([System.ComponentModel.BackgroundWorker]$Worker)
    $script:bgWorker = $Worker

    # Trim log entries older than 90 days
    Trim-LogFile

    $cfg = Load-AppConfig
    if (-not $cfg) { Write-SyncLog "ERROR: config.json not found."; return @{ Success=$false; Stats=@{Saved=0;Skipped=0;Failed=0} } }

    $baseUrl    = ($cfg.apiEndpoint).TrimEnd("/")
    $primKey    = $cfg.primaryKey
    $secKey     = $cfg.secretKey
    $location   = $cfg.location
    $locDesc    = if ($cfg.locationDesc)  { $cfg.locationDesc }  else { $location }
    $client     = if ($cfg.clientName)    { $cfg.clientName }    else { "Unknown" }
    $srcFolder  = $cfg.sourceFolder
    $batchSz    = if ($cfg.batchSize -gt 0) { [int]$cfg.batchSize } else { 80 }
    $licKey     = $cfg.licenseKey
    $dataSource = if ($cfg.dataSource)    { $cfg.dataSource }    else { "database" }
    $dbPath     = $cfg.databasePath
    $fbUser     = if ($cfg.firebird -and $cfg.firebird.user)     { $cfg.firebird.user }     else { "SYSDBA" }
    $fbPass     = if ($cfg.firebird -and $cfg.firebird.password) { $cfg.firebird.password } else { "masterkey" }
    $fbLib      = if ($cfg.firebird -and $cfg.firebird.clientLibrary) { $cfg.firebird.clientLibrary } else { "" }
    $syncDays   = if ($cfg.syncDays -gt 0) { [int]$cfg.syncDays } else { 2 }

    Add-Content -Path $script:logFile -Value "" -Encoding UTF8
    Add-Content -Path $script:logFile -Value "========================================" -Encoding UTF8
    Write-SyncLog "EntryPass-MiHCM Sync"
    Write-SyncLog "Client   : $client"
    Write-SyncLog "Location : $location ($locDesc)"
    if ($dataSource -eq "database") {
        Write-SyncLog "Source   : Database Direct ($dbPath)"
    } else {
        Write-SyncLog "Source   : $srcFolder"
    }
    Write-SyncLog "Endpoint : $baseUrl"

    $licOk = Test-LicenseKey -Key $licKey
    if (-not $licOk) {
        Write-SyncLog "LICENSE: Invalid or expired. Aborting."
        return @{ Success=$false; Stats=@{Saved=0;Skipped=0;Failed=0} }
    }

    $token = Get-MiHCMToken -BaseUrl $baseUrl -PrimaryKey $primKey -SecretKey $secKey
    if (-not $token) {
        Write-SyncLog "Cannot get token -- check API keys."
        return @{ Success=$false; Stats=@{Saved=0;Skipped=0;Failed=0} }
    }

    $allRecords = @()

    if ($dataSource -eq "database") {
        # ---- DATABASE DIRECT MODE ----
        $allRecords = Read-FirebirdDatabase -DbPath $dbPath -FbUser $fbUser -FbPassword $fbPass -ClientLibrary $fbLib -LocationCode $location -SyncDays $syncDays
        if ($null -eq $allRecords) {
            Write-SyncLog "Database read failed. Aborting."
            return @{ Success=$false; Stats=@{Saved=0;Skipped=0;Failed=0} }
        }
    } else {
        # ---- FILE-BASED MODE (original behavior) ----
        $files = Get-ChildItem -Path $srcFolder -Filter "DATA*.txt" -ErrorAction SilentlyContinue
        if (-not $files -or $files.Count -eq 0) {
            Write-SyncLog "No DATA*.txt files in $srcFolder."
            return @{ Success=$true; Stats=@{Saved=0;Skipped=0;Failed=0} }
        }
        Write-SyncLog "Found $($files.Count) file(s)."
        foreach ($file in $files) {
            $recs = Convert-EntryPassFile -InputFile $file.FullName -LocationCode $location
            if ($recs) { $allRecords += $recs }
            Write-SyncLog ""
        }
    }

    if ($allRecords.Count -eq 0) {
        Write-SyncLog "No records to upload."
        return @{ Success=$true; Stats=@{Saved=0;Skipped=0;Failed=0} }
    }

    Write-SyncLog "Total records: $($allRecords.Count)"
    $stats = Upload-Records -BaseUrl $baseUrl -PrimaryKey $primKey -SecretKey $secKey -Token $token -Records $allRecords -BatchSize $batchSz

    if ($dataSource -ne "database") {
        # Only delete source files for file-based mode
        if ($stats.Failed -eq 0) {
            foreach ($file in $files) {
                try { Remove-Item $file.FullName -Force; Write-SyncLog "Deleted: $($file.Name)" } catch { Write-SyncLog "Cannot delete $($file.Name): $_" }
            }
        } else {
            Write-SyncLog "WARNING: $($stats.Failed) failed -- source files NOT deleted."
        }
    }

    $result = if ($stats.Failed -eq 0 -and $stats.Saved -gt 0) { "SUCCESS" }
              elseif ($stats.Failed -gt 0 -and $stats.Saved -gt 0) { "PARTIAL" }
              elseif ($stats.Saved -eq 0 -and $stats.Skipped -gt 0) { "ALL DUPLICATES" }
              else { "FAILURE" }

    Add-Content -Path $script:logFile -Value "" -Encoding UTF8
    Write-SyncLog "========================================"
    Write-SyncLog "SUMMARY: $result | Saved:$($stats.Saved) Skipped:$($stats.Skipped) Failed:$($stats.Failed)"
    Write-SyncLog "========================================"

    $script:bgWorker = $null
    return @{ Success=$true; Stats=$stats; Result=$result }
}

# ============================================================
# TASK SCHEDULER HELPER
# ============================================================
function Install-SyncTask {
    param([string]$Location,[string]$Frequency)
    $taskName  = "EntryPass-MiHCM Sync - $Location"
    $psExe     = "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    $script    = Join-Path $script:appDir "EntryPassSync.ps1"
    $args      = "-ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File `"$script`" -RunSyncOnly"

    $interval  = switch ($Frequency) {
        "15min"  { "PT15M" }
        "30min"  { "PT30M" }
        "1hour"  { "PT1H" }
        "2hour"  { "PT2H" }
        default  { "PT30M" }
    }

    $startBoundary = (Get-Date -Format "yyyy-MM-ddT06:00:00")
    $taskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo><Description>EntryPass MiHCM sync for $Location</Description></RegistrationInfo>
  <Triggers>
    <TimeTrigger>
      <Repetition><Interval>$interval</Interval><Duration>P1D</Duration><StopAtDurationEnd>false</StopAtDurationEnd></Repetition>
      <StartBoundary>$startBoundary</StartBoundary>
      <Enabled>true</Enabled>
    </TimeTrigger>
  </Triggers>
  <Principals><Principal id="Author"><UserId>S-1-5-18</UserId><LogonType>Password</LogonType><RunLevel>HighestAvailable</RunLevel></Principal></Principals>
  <Settings><MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy><DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries><StopIfGoingOnBatteries>false</StopIfGoingOnBatteries><ExecutionTimeLimit>PT1H</ExecutionTimeLimit><Priority>7</Priority><StartWhenAvailable>true</StartWhenAvailable><RestartOnFailure><Interval>PT5M</Interval><Count>3</Count></RestartOnFailure></Settings>
  <Actions><Exec><Command>$psExe</Command><Arguments>$args</Arguments><WorkingDirectory>$($script:appDir)</WorkingDirectory></Exec></Actions>
</Task>
"@
    $tmpXml = Join-Path $env:TEMP "ep_sync_task.xml"
    [System.IO.File]::WriteAllText($tmpXml, $taskXml, [System.Text.Encoding]::Unicode)
    try {
        $out = & schtasks.exe /Create /TN $taskName /XML $tmpXml /F 2>&1
        Remove-Item $tmpXml -Force -ErrorAction SilentlyContinue
        if ($LASTEXITCODE -eq 0) { return @{Success=$true;Message="Task '$taskName' created."} }
        return @{Success=$false;Message="schtasks error (code $LASTEXITCODE): $out"}
    } catch {
        Remove-Item $tmpXml -Force -ErrorAction SilentlyContinue
        return @{Success=$false;Message="Exception: $_"}
    }
}

function Remove-SyncTask {
    param([string]$Location)
    $taskName = "EntryPass-MiHCM Sync - $Location"
    $out = & schtasks.exe /Delete /TN $taskName /F 2>&1
    return $LASTEXITCODE -eq 0
}

# ============================================================
# HEADLESS SYNC MODE (called by Task Scheduler)
# ============================================================
if ($args -contains "-RunSyncOnly") {
    $dummy = [System.ComponentModel.BackgroundWorker]::new()
    Run-FullSync -Worker $dummy | Out-Null
    exit 0
}

# ============================================================
# =================== BUILD THE MAIN FORM ===================
# ============================================================
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.SuspendLayout()
$mainForm.Text            = "EntryPass-MiHCM Sync v1.0"
$mainForm.Size            = New-Object System.Drawing.Size(900,620)
$mainForm.MinimumSize     = New-Object System.Drawing.Size(900,620)
$mainForm.StartPosition   = "CenterScreen"
$mainForm.FormBorderStyle = "Sizable"
$mainForm.MaximizeBox     = $true
$mainForm.BackColor       = $clrPanelBg
$mainForm.Font            = New-Object System.Drawing.Font("Segoe UI",9)

# ============================================================
# LEFT SIDEBAR (200px wide)
# ============================================================
$sidebar = New-Object System.Windows.Forms.Panel
$sidebar.Size      = New-Object System.Drawing.Size(200,620)
$sidebar.Location  = New-Object System.Drawing.Point(0,0)
$sidebar.BackColor = $clrSidebar
$sidebar.Anchor    = "Top,Left,Bottom"
$mainForm.Controls.Add($sidebar)

# App title in sidebar
$sideTitle = New-Object System.Windows.Forms.Label
$sideTitle.Text      = "EntryPass Sync"
$sideTitle.Font      = New-Object System.Drawing.Font("Segoe UI",11,[System.Drawing.FontStyle]::Bold)
$sideTitle.ForeColor = [System.Drawing.Color]::White
$sideTitle.Location  = New-Object System.Drawing.Point(16,20)
$sideTitle.Size      = New-Object System.Drawing.Size(168,26)
$sidebar.Controls.Add($sideTitle)

$sideSubtitle = New-Object System.Windows.Forms.Label
$sideSubtitle.Text      = "MiHCM Integration"
$sideSubtitle.Font      = New-Object System.Drawing.Font("Segoe UI",8)
$sideSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(140,160,190)
$sideSubtitle.Location  = New-Object System.Drawing.Point(16,46)
$sideSubtitle.Size      = New-Object System.Drawing.Size(168,20)
$sidebar.Controls.Add($sideSubtitle)

$sideDiv = New-Object System.Windows.Forms.Label
$sideDiv.BackColor = [System.Drawing.Color]::FromArgb(60,80,110)
$sideDiv.Location  = New-Object System.Drawing.Point(16,72)
$sideDiv.Size      = New-Object System.Drawing.Size(168,1)
$sidebar.Controls.Add($sideDiv)

# Navigation items definition
$navDefs = @(
    @{Key="navDashboard";    Text="  Dashboard";    Y=86 },
    @{Key="navConfig";       Text="  Configuration"; Y=128},
    @{Key="navLogs";         Text="  Logs";          Y=170},
    @{Key="navAbout";        Text="  About";         Y=212}
)

$script:navItems  = @{}
$script:contentPanels = @{}

foreach ($nd in $navDefs) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = $nd.Text
    $lbl.Font      = New-Object System.Drawing.Font("Segoe UI",9.5)
    $lbl.ForeColor = [System.Drawing.Color]::FromArgb(200,210,225)
    $lbl.BackColor = $clrSidebar
    $lbl.Location  = New-Object System.Drawing.Point(0,$nd.Y)
    $lbl.Size      = New-Object System.Drawing.Size(200,36)
    $lbl.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $lbl.Tag       = $nd.Key
    $sidebar.Controls.Add($lbl)
    $script:navItems[$nd.Key] = $lbl
}

# Sidebar footer
$sideFooter1 = New-Object System.Windows.Forms.Label
$sideFooter1.Text      = "Dajayana Trading"
$sideFooter1.Font      = New-Object System.Drawing.Font("Segoe UI",8,[System.Drawing.FontStyle]::Bold)
$sideFooter1.ForeColor = [System.Drawing.Color]::FromArgb(140,160,190)
$sideFooter1.Location  = New-Object System.Drawing.Point(12,560)
$sideFooter1.Size      = New-Object System.Drawing.Size(176,18)
$sideFooter1.Anchor    = "Left,Bottom"
$sidebar.Controls.Add($sideFooter1)

$sideFooter2 = New-Object System.Windows.Forms.Label
$sideFooter2.Text      = "www.dajayana.com"
$sideFooter2.Font      = New-Object System.Drawing.Font("Segoe UI",7.5)
$sideFooter2.ForeColor = [System.Drawing.Color]::FromArgb(100,120,150)
$sideFooter2.Location  = New-Object System.Drawing.Point(12,578)
$sideFooter2.Size      = New-Object System.Drawing.Size(176,18)
$sideFooter2.Anchor    = "Left,Bottom"
$sidebar.Controls.Add($sideFooter2)

# ============================================================
# CONTENT AREA (right of sidebar, 700x620)
# ============================================================
$contentArea = New-Object System.Windows.Forms.Panel
$contentArea.Size      = New-Object System.Drawing.Size(700,620)
$contentArea.Location  = New-Object System.Drawing.Point(200,0)
$contentArea.BackColor = $clrPanelBg
$contentArea.Anchor    = "Top,Left,Right,Bottom"
$contentArea.BorderStyle = "FixedSingle"
$mainForm.Controls.Add($contentArea)

# Helper to make a content panel
function New-ContentPanel {
    $p = New-Object System.Windows.Forms.Panel
    $p.Dock      = "Fill"
    $p.BackColor = $clrPanelBg
    $p.Visible   = $false
    $p.AutoScroll = $true
    $contentArea.Controls.Add($p)
    return $p
}

# ============================================================
# DASHBOARD PANEL
# ============================================================
$panelDashboard = New-ContentPanel

$dashTitle = New-Object System.Windows.Forms.Label
$dashTitle.Text      = "Dashboard"
$dashTitle.Font      = New-Object System.Drawing.Font("Segoe UI",16,[System.Drawing.FontStyle]::Bold)
$dashTitle.ForeColor = $clrTextDark
$dashTitle.Location  = New-Object System.Drawing.Point(24,20)
$dashTitle.Size      = New-Object System.Drawing.Size(400,36)
$panelDashboard.Controls.Add($dashTitle)

# Status card
$statusCard = New-Object System.Windows.Forms.Panel
$statusCard.BackColor = [System.Drawing.Color]::White
$statusCard.Location  = New-Object System.Drawing.Point(24,64)
$statusCard.Size      = New-Object System.Drawing.Size(652,78)
$statusCard.Anchor     = "Top,Left,Right"
$panelDashboard.Controls.Add($statusCard)

$script:lblStatus = New-Object System.Windows.Forms.Label
$script:lblStatus.Text      = "Not Configured"
$script:lblStatus.Font      = New-Object System.Drawing.Font("Segoe UI",13,[System.Drawing.FontStyle]::Bold)
$script:lblStatus.ForeColor = $clrGrey
$script:lblStatus.Location  = New-Object System.Drawing.Point(14,16)
$script:lblStatus.Size      = New-Object System.Drawing.Size(450,38)
$statusCard.Controls.Add($script:lblStatus)

$script:lblLastSync = New-Object System.Windows.Forms.Label
$script:lblLastSync.Text      = "Last sync: Never"
$script:lblLastSync.Font      = New-Object System.Drawing.Font("Segoe UI",8)
$script:lblLastSync.ForeColor = $clrTextDim
$script:lblLastSync.Location  = New-Object System.Drawing.Point(440,42)
$script:lblLastSync.Size      = New-Object System.Drawing.Size(200,18)
$statusCard.Controls.Add($script:lblLastSync)

# Stats row
$statsCard = New-Object System.Windows.Forms.Panel
$statsCard.BackColor = [System.Drawing.Color]::White
$statsCard.Location  = New-Object System.Drawing.Point(24,152)
$statsCard.Size      = New-Object System.Drawing.Size(652,72)
$statsCard.Anchor     = "Top,Left,Right"
$panelDashboard.Controls.Add($statsCard)

function Make-StatLabel {
    param([string]$Title,[int]$X,[System.Drawing.Color]$Color)
    $t = New-Object System.Windows.Forms.Label
    $t.Text      = $Title
    $t.Font      = New-Object System.Drawing.Font("Segoe UI",7.5)
    $t.ForeColor = $clrTextDim
    $t.Location  = New-Object System.Drawing.Point($X,10)
    $t.Size      = New-Object System.Drawing.Size(120,16)
    $statsCard.Controls.Add($t)
    $v = New-Object System.Windows.Forms.Label
    $v.Text      = "0"
    $v.Font      = New-Object System.Drawing.Font("Segoe UI",18,[System.Drawing.FontStyle]::Bold)
    $v.ForeColor = $Color
    $v.Location  = New-Object System.Drawing.Point($X,26)
    $v.Size      = New-Object System.Drawing.Size(120,38)
    $statsCard.Controls.Add($v)
    return $v
}
$script:lblSaved   = Make-StatLabel "Uploaded"  14  $clrGreen
$script:lblSkipped = New-Object System.Windows.Forms.Label  # hidden - kept for internal tracking
$script:lblFailed  = Make-StatLabel "Failed"    230 ([System.Drawing.Color]::FromArgb(210,60,60))

# Site info card
$siteCard = New-Object System.Windows.Forms.Panel
$siteCard.BackColor = [System.Drawing.Color]::White
$siteCard.Location  = New-Object System.Drawing.Point(24,234)
$siteCard.Size      = New-Object System.Drawing.Size(652,92)
$siteCard.Anchor     = "Top,Left,Right"
$panelDashboard.Controls.Add($siteCard)

$lblSiteTitle = New-Object System.Windows.Forms.Label
$lblSiteTitle.Text      = "Site Information"
$lblSiteTitle.Font      = New-Object System.Drawing.Font("Segoe UI",8,[System.Drawing.FontStyle]::Bold)
$lblSiteTitle.ForeColor = $clrTextDim
$lblSiteTitle.Location  = New-Object System.Drawing.Point(14,8)
$lblSiteTitle.Size      = New-Object System.Drawing.Size(300,18)
$siteCard.Controls.Add($lblSiteTitle)

$script:lblSiteClient   = New-Object System.Windows.Forms.Label
$script:lblSiteLocation = New-Object System.Windows.Forms.Label
$script:lblSiteFolder   = New-Object System.Windows.Forms.Label
$locs = @(
    @{Lbl=$script:lblSiteClient;   Text="Client:";   Y=28 },
    @{Lbl=$script:lblSiteLocation; Text="Location:"; Y=48 },
    @{Lbl=$script:lblSiteFolder;   Text="Source:";   Y=68 }
)
foreach ($li in $locs) {
    $hl = New-Object System.Windows.Forms.Label
    $hl.Text      = $li.Text
    $hl.Font      = New-Object System.Drawing.Font("Segoe UI",8)
    $hl.ForeColor = $clrTextDim
    $hl.Location  = New-Object System.Drawing.Point(14,$li.Y)
    $hl.Size      = New-Object System.Drawing.Size(80,18)
    $siteCard.Controls.Add($hl)
    $li.Lbl.Text      = "--"
    $li.Lbl.Font      = New-Object System.Drawing.Font("Segoe UI",8)
    $li.Lbl.ForeColor = $clrTextDark
    $li.Lbl.Location  = New-Object System.Drawing.Point(100,$li.Y)
    $li.Lbl.Size      = New-Object System.Drawing.Size(530,18)
    $siteCard.Controls.Add($li.Lbl)
}

# Sync Now button
$script:btnSyncNow = New-Object System.Windows.Forms.Button
$script:btnSyncNow.Text      = "Sync Now"
$script:btnSyncNow.Font      = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$script:btnSyncNow.BackColor = $clrBlue
$script:btnSyncNow.ForeColor = [System.Drawing.Color]::White
$script:btnSyncNow.FlatStyle = "Flat"
$script:btnSyncNow.FlatAppearance.BorderSize = 0
$script:btnSyncNow.Location  = New-Object System.Drawing.Point(24,336)
$script:btnSyncNow.Size      = New-Object System.Drawing.Size(140,38)
$script:btnSyncNow.Cursor    = [System.Windows.Forms.Cursors]::Hand
$panelDashboard.Controls.Add($script:btnSyncNow)

$script:lblSyncStatus = New-Object System.Windows.Forms.Label
$script:lblSyncStatus.Text      = ""
$script:lblSyncStatus.Font      = New-Object System.Drawing.Font("Segoe UI",8)
$script:lblSyncStatus.ForeColor = $clrTextDim
$script:lblSyncStatus.Location  = New-Object System.Drawing.Point(176,348)
$script:lblSyncStatus.Size      = New-Object System.Drawing.Size(400,18)
$panelDashboard.Controls.Add($script:lblSyncStatus)

# Live log box
$lblLiveLog = New-Object System.Windows.Forms.Label
$lblLiveLog.Text      = "Live Output"
$lblLiveLog.Anchor    = "Top,Left"
$lblLiveLog.Font      = New-Object System.Drawing.Font("Segoe UI",8)
$lblLiveLog.ForeColor = $clrTextDim
$lblLiveLog.Location  = New-Object System.Drawing.Point(24,382)
$lblLiveLog.Size      = New-Object System.Drawing.Size(200,18)
$panelDashboard.Controls.Add($lblLiveLog)

$script:txtLiveLog = New-Object System.Windows.Forms.TextBox
$script:txtLiveLog.Multiline    = $true
$script:txtLiveLog.ScrollBars   = "Vertical"
$script:txtLiveLog.ReadOnly     = $true
$script:txtLiveLog.BackColor    = $clrDarkBox
$script:txtLiveLog.ForeColor    = [System.Drawing.Color]::FromArgb(180,220,180)
$script:txtLiveLog.Font         = New-Object System.Drawing.Font("Consolas",8)
$script:txtLiveLog.Location     = New-Object System.Drawing.Point(24,400)
$script:txtLiveLog.Size         = New-Object System.Drawing.Size(652,150)
$script:txtLiveLog.BorderStyle  = "None"
$script:txtLiveLog.Anchor       = "Top,Left,Right"
$panelDashboard.Controls.Add($script:txtLiveLog)

# ============================================================
# CONFIGURATION PANEL
# ============================================================
$panelConfig = New-ContentPanel

$cfgTitle = New-Object System.Windows.Forms.Label
$cfgTitle.Text      = "Configuration"
$cfgTitle.Font      = New-Object System.Drawing.Font("Segoe UI",16,[System.Drawing.FontStyle]::Bold)
$cfgTitle.ForeColor = $clrTextDark
$cfgTitle.Location  = New-Object System.Drawing.Point(24,20)
$cfgTitle.Size      = New-Object System.Drawing.Size(400,36)
$panelConfig.Controls.Add($cfgTitle)

# Config fields container (scrollable)
$cfgScroll = New-Object System.Windows.Forms.Panel
$cfgScroll.AutoScroll = $true
$cfgScroll.Location   = New-Object System.Drawing.Point(24,64)
$cfgScroll.Size       = New-Object System.Drawing.Size(652,470)
$cfgScroll.BackColor  = $clrPanelBg
$panelConfig.Controls.Add($cfgScroll)

$cfgY = 8

function Add-CfgField {
    param([string]$Label,[System.Windows.Forms.Control]$Control,[int]$Height=28)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text      = $Label
    $lbl.Font      = New-Object System.Drawing.Font("Segoe UI",8.5)
    $lbl.ForeColor = $clrTextDim
    $lbl.Location  = New-Object System.Drawing.Point(4,$cfgY)
    $lbl.Size      = New-Object System.Drawing.Size(175,20)
    $cfgScroll.Controls.Add($lbl)
    $Control.Location = New-Object System.Drawing.Point(183,$cfgY)
    $Control.Size     = New-Object System.Drawing.Size(440,$Height)
    $Control.Font     = New-Object System.Drawing.Font("Segoe UI",9)
    $cfgScroll.Controls.Add($Control)
    $script:cfgY += ($Height + 10)
}

# Hidden config fields (pre-configured by Dajayana, not shown to site IT)
$script:txtCfgLicense = New-Object System.Windows.Forms.TextBox
$script:btnValidateLicense = New-Object System.Windows.Forms.Button
$script:txtCfgPrimary = New-Object System.Windows.Forms.TextBox
$script:txtCfgSecret = New-Object System.Windows.Forms.TextBox
$script:txtCfgSecret.UseSystemPasswordChar = $true
$script:cmbCfgEndpoint = New-Object System.Windows.Forms.ComboBox
$script:cmbCfgEndpoint.DropDownStyle = "DropDownList"
[void]$script:cmbCfgEndpoint.Items.Add("Production (api.mihcm.com)")
[void]$script:cmbCfgEndpoint.Items.Add("UAT (api.mihcm.com/uat)")
$script:cmbCfgEndpoint.SelectedIndex = 0
# These controls exist for config load/save but are NOT added to the UI

$script:txtCfgClient = New-Object System.Windows.Forms.TextBox
Add-CfgField "Client Name *" $script:txtCfgClient

$script:txtCfgLocation = New-Object System.Windows.Forms.TextBox
Add-CfgField "Location Code *" $script:txtCfgLocation

$script:txtCfgLocationDesc = New-Object System.Windows.Forms.TextBox
Add-CfgField "Location Description" $script:txtCfgLocationDesc

# ---- DATA SOURCE MODE SECTION ----
$sepDsrc = New-Object System.Windows.Forms.Label
$sepDsrc.BackColor = [System.Drawing.Color]::FromArgb(210,215,225)
$sepDsrc.Location  = New-Object System.Drawing.Point(4,$cfgY)
$sepDsrc.Size      = New-Object System.Drawing.Size(619,1)
$cfgScroll.Controls.Add($sepDsrc)
$cfgY += 8

$lblDsrcHdr = New-Object System.Windows.Forms.Label
$lblDsrcHdr.Text      = "Data Source Mode"
$lblDsrcHdr.Font      = New-Object System.Drawing.Font("Segoe UI",8.5,[System.Drawing.FontStyle]::Bold)
$lblDsrcHdr.ForeColor = $clrTextDim
$lblDsrcHdr.Location  = New-Object System.Drawing.Point(4,$cfgY)
$lblDsrcHdr.Size      = New-Object System.Drawing.Size(619,20)
$cfgScroll.Controls.Add($lblDsrcHdr)
$cfgY += 26

# Radio: Database Direct
$script:rdoDbMode = New-Object System.Windows.Forms.RadioButton
$script:rdoDbMode.Text     = "Database Direct (Firebird) -- reads attendance directly from TRANS.FDB"
$script:rdoDbMode.Font     = New-Object System.Drawing.Font("Segoe UI",9)
$script:rdoDbMode.Checked  = $true
$script:rdoDbMode.Location = New-Object System.Drawing.Point(4,$cfgY)
$script:rdoDbMode.Size     = New-Object System.Drawing.Size(619,22)
$cfgScroll.Controls.Add($script:rdoDbMode)
$cfgY += 26

# INDENTED: Database path row
$script:lblDbPath = New-Object System.Windows.Forms.Label
$script:lblDbPath.Text      = "Database Path"
$script:lblDbPath.Font      = New-Object System.Drawing.Font("Segoe UI",8.5)
$script:lblDbPath.ForeColor = $clrTextDim
$script:lblDbPath.Location  = New-Object System.Drawing.Point(30,$cfgY)
$script:lblDbPath.Size      = New-Object System.Drawing.Size(149,20)
$cfgScroll.Controls.Add($script:lblDbPath)

$script:txtCfgDbPath = New-Object System.Windows.Forms.TextBox
$script:txtCfgDbPath.Location = New-Object System.Drawing.Point(183,$cfgY)
$script:txtCfgDbPath.Size     = New-Object System.Drawing.Size(360,28)
$script:txtCfgDbPath.Font     = New-Object System.Drawing.Font("Segoe UI",9)
$cfgScroll.Controls.Add($script:txtCfgDbPath)

$script:btnBrowseDbPath = New-Object System.Windows.Forms.Button
$script:btnBrowseDbPath.Text      = "Browse"
$script:btnBrowseDbPath.Font      = New-Object System.Drawing.Font("Segoe UI",8.5)
$script:btnBrowseDbPath.FlatStyle = "Flat"
$script:btnBrowseDbPath.BackColor = [System.Drawing.Color]::FromArgb(230,235,240)
$script:btnBrowseDbPath.Location  = New-Object System.Drawing.Point(549,$cfgY)
$script:btnBrowseDbPath.Size      = New-Object System.Drawing.Size(74,28)
$script:btnBrowseDbPath.Cursor    = [System.Windows.Forms.Cursors]::Hand
$cfgScroll.Controls.Add($script:btnBrowseDbPath)
$cfgY += 38

# Firebird credentials - hidden from UI, pre-configured in config.json
$script:lblFbUser = New-Object System.Windows.Forms.Label
$script:txtCfgFbUser = New-Object System.Windows.Forms.TextBox
$script:txtCfgFbUser.Text = "SYSDBA"
$script:lblFbPass = New-Object System.Windows.Forms.Label
$script:txtCfgFbPass = New-Object System.Windows.Forms.TextBox
$script:txtCfgFbPass.Text = "masterkey"
$script:txtCfgFbPass.UseSystemPasswordChar = $true
$script:lblFbLib = New-Object System.Windows.Forms.Label
$script:txtCfgFbLib = New-Object System.Windows.Forms.TextBox

# INDENTED: Sync days spinner
$script:lblSyncDays = New-Object System.Windows.Forms.Label
$script:lblSyncDays.Text      = "Sync Days (DB mode)"
$script:lblSyncDays.Font      = New-Object System.Drawing.Font("Segoe UI",8.5)
$script:lblSyncDays.ForeColor = $clrTextDim
$script:lblSyncDays.Location  = New-Object System.Drawing.Point(30,$cfgY)
$script:lblSyncDays.Size      = New-Object System.Drawing.Size(149,20)
$cfgScroll.Controls.Add($script:lblSyncDays)
$script:numSyncDays = New-Object System.Windows.Forms.NumericUpDown
$script:numSyncDays.Minimum       = 1
$script:numSyncDays.Maximum       = 30
$script:numSyncDays.Value         = 1
$script:numSyncDays.DecimalPlaces = 0
$script:numSyncDays.Location      = New-Object System.Drawing.Point(183,$cfgY)
$script:numSyncDays.Size          = New-Object System.Drawing.Size(440,28)
$script:numSyncDays.Font          = New-Object System.Drawing.Font("Segoe UI",9)
$cfgScroll.Controls.Add($script:numSyncDays)
$cfgY += 38

# INDENTED: Test DB Connection button + status label
$script:btnTestDbConn = New-Object System.Windows.Forms.Button
$script:btnTestDbConn.Text      = "Test DB Connection"
$script:btnTestDbConn.Font      = New-Object System.Drawing.Font("Segoe UI",8.5)
$script:btnTestDbConn.FlatStyle = "Flat"
$script:btnTestDbConn.BackColor = [System.Drawing.Color]::FromArgb(230,235,240)
$script:btnTestDbConn.Location  = New-Object System.Drawing.Point(30,$cfgY)
$script:btnTestDbConn.Size      = New-Object System.Drawing.Size(160,28)
$script:btnTestDbConn.Cursor    = [System.Windows.Forms.Cursors]::Hand
$cfgScroll.Controls.Add($script:btnTestDbConn)

$script:lblDbConnStatus = New-Object System.Windows.Forms.Label
$script:lblDbConnStatus.Text      = ""
$script:lblDbConnStatus.Font      = New-Object System.Drawing.Font("Segoe UI",8.5)
$script:lblDbConnStatus.ForeColor = $clrTextDim
$script:lblDbConnStatus.Location  = New-Object System.Drawing.Point(198,$cfgY)
$script:lblDbConnStatus.Size      = New-Object System.Drawing.Size(425,28)
$cfgScroll.Controls.Add($script:lblDbConnStatus)
$cfgY += 36

# Radio: File-based
$script:rdoFileMode = New-Object System.Windows.Forms.RadioButton
$script:rdoFileMode.Text     = "File-based (DATA*.txt) -- reads exported text files from EntryPass"
$script:rdoFileMode.Font     = New-Object System.Drawing.Font("Segoe UI",9)
$script:rdoFileMode.Checked  = $false
$script:rdoFileMode.Location = New-Object System.Drawing.Point(4,$cfgY)
$script:rdoFileMode.Size     = New-Object System.Drawing.Size(619,22)
$cfgScroll.Controls.Add($script:rdoFileMode)
$cfgY += 26

# INDENTED: Source folder row with Browse button
$script:lblSrcF = New-Object System.Windows.Forms.Label
$script:lblSrcF.Text      = "Source Folder *"
$script:lblSrcF.Font      = New-Object System.Drawing.Font("Segoe UI",8.5)
$script:lblSrcF.ForeColor = $clrTextDim
$script:lblSrcF.Location  = New-Object System.Drawing.Point(30,$cfgY)
$script:lblSrcF.Size      = New-Object System.Drawing.Size(149,20)
$cfgScroll.Controls.Add($script:lblSrcF)

$script:txtCfgSourceFolder = New-Object System.Windows.Forms.TextBox
$script:txtCfgSourceFolder.Location = New-Object System.Drawing.Point(183,$cfgY)
$script:txtCfgSourceFolder.Size     = New-Object System.Drawing.Size(360,28)
$script:txtCfgSourceFolder.Font     = New-Object System.Drawing.Font("Segoe UI",9)
$cfgScroll.Controls.Add($script:txtCfgSourceFolder)

$script:btnBrowseFolder = New-Object System.Windows.Forms.Button
$script:btnBrowseFolder.Text      = "Browse"
$script:btnBrowseFolder.Font      = New-Object System.Drawing.Font("Segoe UI",8.5)
$script:btnBrowseFolder.FlatStyle = "Flat"
$script:btnBrowseFolder.BackColor = [System.Drawing.Color]::FromArgb(230,235,240)
$script:btnBrowseFolder.Location  = New-Object System.Drawing.Point(549,$cfgY)
$script:btnBrowseFolder.Size      = New-Object System.Drawing.Size(74,28)
$script:btnBrowseFolder.Cursor    = [System.Windows.Forms.Cursors]::Hand
$cfgScroll.Controls.Add($script:btnBrowseFolder)
$cfgY += 32

$script:btnSyncFilesNow = New-Object System.Windows.Forms.Button
$script:btnSyncFilesNow.Text      = "Sync Files Now"
$script:btnSyncFilesNow.Font      = New-Object System.Drawing.Font("Segoe UI",8.5,[System.Drawing.FontStyle]::Bold)
$script:btnSyncFilesNow.FlatStyle = "Flat"
$script:btnSyncFilesNow.BackColor = [System.Drawing.Color]::FromArgb(0,51,102)
$script:btnSyncFilesNow.ForeColor = [System.Drawing.Color]::White
$script:btnSyncFilesNow.Location  = New-Object System.Drawing.Point(30,$cfgY)
$script:btnSyncFilesNow.Size      = New-Object System.Drawing.Size(120,28)
$script:btnSyncFilesNow.Cursor    = [System.Windows.Forms.Cursors]::Hand
$cfgScroll.Controls.Add($script:btnSyncFilesNow)
$cfgY += 38

# Helper to enable/disable database vs file-mode controls with visual greying
function Update-DataSourceUI {
    $dbMode   = $script:rdoDbMode.Checked
    $dimColor = [System.Drawing.Color]::FromArgb(190,195,205)
    $normColor = $clrTextDim
    # DB mode controls
    $script:txtCfgDbPath.Enabled    = $dbMode
    $script:btnBrowseDbPath.Enabled = $dbMode
    $script:numSyncDays.Enabled     = $dbMode
    $script:btnTestDbConn.Enabled   = $dbMode
    $script:lblDbConnStatus.Text    = ""
    $script:lblDbPath.ForeColor     = if ($dbMode) { $normColor } else { $dimColor }
    $script:lblSyncDays.ForeColor   = if ($dbMode) { $normColor } else { $dimColor }
    # File mode controls
    $script:txtCfgSourceFolder.Enabled = (-not $dbMode)
    $script:btnBrowseFolder.Enabled    = (-not $dbMode)
    $script:btnSyncFilesNow.Enabled    = (-not $dbMode)
    $script:lblSrcF.ForeColor          = if (-not $dbMode) { $normColor } else { $dimColor }
}

$script:rdoFileMode.add_CheckedChanged({ Update-DataSourceUI })
$script:rdoDbMode.add_CheckedChanged({  Update-DataSourceUI })

# Schedule - hidden (now uses built-in 15-min timer, not Task Scheduler)
$script:chkSchedule = New-Object System.Windows.Forms.CheckBox
$script:chkSchedule.Checked = $true
$script:cmbFrequency = New-Object System.Windows.Forms.ComboBox
[void]$script:cmbFrequency.Items.Add("Every 15 minutes")
$script:cmbFrequency.SelectedIndex = 0

$sepCfg2 = New-Object System.Windows.Forms.Label
$sepCfg2.BackColor = [System.Drawing.Color]::FromArgb(210,215,225)
$sepCfg2.Location  = New-Object System.Drawing.Point(4,$cfgY)
$sepCfg2.Size      = New-Object System.Drawing.Size(619,1)
$cfgScroll.Controls.Add($sepCfg2)
$cfgY += 10

# Config status label
$script:lblCfgStatus = New-Object System.Windows.Forms.Label
$script:lblCfgStatus.Text      = ""
$script:lblCfgStatus.Font      = New-Object System.Drawing.Font("Segoe UI",8.5)
$script:lblCfgStatus.ForeColor = $clrTextDim
$script:lblCfgStatus.Location  = New-Object System.Drawing.Point(4,$cfgY)
$script:lblCfgStatus.Size      = New-Object System.Drawing.Size(619,22)
$cfgScroll.Controls.Add($script:lblCfgStatus)
$cfgY += 30

# Config action buttons
$script:btnTestConn = New-Object System.Windows.Forms.Button
$script:btnTestConn.Text      = "Test API Connection"
$script:btnTestConn.Font      = New-Object System.Drawing.Font("Segoe UI",8.5)
$script:btnTestConn.FlatStyle = "Flat"
$script:btnTestConn.BackColor = [System.Drawing.Color]::FromArgb(230,235,240)
$script:btnTestConn.Location  = New-Object System.Drawing.Point(4,$cfgY)
$script:btnTestConn.Size      = New-Object System.Drawing.Size(140,30)
$script:btnTestConn.Cursor    = [System.Windows.Forms.Cursors]::Hand
$cfgScroll.Controls.Add($script:btnTestConn)

$script:btnSaveInstall = New-Object System.Windows.Forms.Button
$script:btnSaveInstall.Text      = "Save Configuration"
$script:btnSaveInstall.Font      = New-Object System.Drawing.Font("Segoe UI",8.5,[System.Drawing.FontStyle]::Bold)
$script:btnSaveInstall.FlatStyle = "Flat"
$script:btnSaveInstall.FlatAppearance.BorderSize = 0
$script:btnSaveInstall.BackColor = $clrBlue
$script:btnSaveInstall.ForeColor = [System.Drawing.Color]::White
$script:btnSaveInstall.Location  = New-Object System.Drawing.Point(154,$cfgY)
$script:btnSaveInstall.Size      = New-Object System.Drawing.Size(200,30)
$script:btnSaveInstall.Cursor    = [System.Windows.Forms.Cursors]::Hand
$cfgScroll.Controls.Add($script:btnSaveInstall)

$script:btnSaveOnly = New-Object System.Windows.Forms.Button
$script:btnSaveOnly.Text      = "Save Only"
$script:btnSaveOnly.Font      = New-Object System.Drawing.Font("Segoe UI",8.5)
$script:btnSaveOnly.FlatStyle = "Flat"
$script:btnSaveOnly.BackColor = [System.Drawing.Color]::FromArgb(230,235,240)
$script:btnSaveOnly.Location  = New-Object System.Drawing.Point(364,$cfgY)
$script:btnSaveOnly.Size      = New-Object System.Drawing.Size(120,30)
$script:btnSaveOnly.Cursor    = [System.Windows.Forms.Cursors]::Hand
$cfgScroll.Controls.Add($script:btnSaveOnly)

# ============================================================
# LOGS PANEL
# ============================================================
$panelLogs = New-ContentPanel

$logsTitle = New-Object System.Windows.Forms.Label
$logsTitle.Text      = "Sync Logs"
$logsTitle.Font      = New-Object System.Drawing.Font("Segoe UI",16,[System.Drawing.FontStyle]::Bold)
$logsTitle.ForeColor = $clrTextDark
$logsTitle.Location  = New-Object System.Drawing.Point(24,20)
$logsTitle.Size      = New-Object System.Drawing.Size(400,36)
$panelLogs.Controls.Add($logsTitle)

$script:btnRefreshLog = New-Object System.Windows.Forms.Button
$script:btnRefreshLog.Text      = "Refresh"
$script:btnRefreshLog.Font      = New-Object System.Drawing.Font("Segoe UI",8.5)
$script:btnRefreshLog.FlatStyle = "Flat"
$script:btnRefreshLog.BackColor = [System.Drawing.Color]::FromArgb(230,235,240)
$script:btnRefreshLog.Location  = New-Object System.Drawing.Point(24,64)
$script:btnRefreshLog.Size      = New-Object System.Drawing.Size(90,28)
$script:btnRefreshLog.Cursor    = [System.Windows.Forms.Cursors]::Hand
$panelLogs.Controls.Add($script:btnRefreshLog)

$script:btnClearLog = New-Object System.Windows.Forms.Button
$script:btnClearLog.Text      = "Clear"
$script:btnClearLog.Font      = New-Object System.Drawing.Font("Segoe UI",8.5)
$script:btnClearLog.FlatStyle = "Flat"
$script:btnClearLog.BackColor = [System.Drawing.Color]::FromArgb(230,235,240)
$script:btnClearLog.Location  = New-Object System.Drawing.Point(122,64)
$script:btnClearLog.Size      = New-Object System.Drawing.Size(90,28)
$script:btnClearLog.Cursor    = [System.Windows.Forms.Cursors]::Hand
$panelLogs.Controls.Add($script:btnClearLog)

$script:btnOpenNotepad = New-Object System.Windows.Forms.Button
$script:btnOpenNotepad.Text      = "Open in Notepad"
$script:btnOpenNotepad.Font      = New-Object System.Drawing.Font("Segoe UI",8.5)
$script:btnOpenNotepad.FlatStyle = "Flat"
$script:btnOpenNotepad.BackColor = [System.Drawing.Color]::FromArgb(230,235,240)
$script:btnOpenNotepad.Location  = New-Object System.Drawing.Point(220,64)
$script:btnOpenNotepad.Size      = New-Object System.Drawing.Size(130,28)
$script:btnOpenNotepad.Cursor    = [System.Windows.Forms.Cursors]::Hand
$panelLogs.Controls.Add($script:btnOpenNotepad)

$script:txtLogsView = New-Object System.Windows.Forms.TextBox
$script:txtLogsView.Multiline   = $true
$script:txtLogsView.ScrollBars  = "Both"
$script:txtLogsView.ReadOnly    = $true
$script:txtLogsView.BackColor   = $clrDarkBox
$script:txtLogsView.ForeColor   = [System.Drawing.Color]::FromArgb(180,220,180)
$script:txtLogsView.Font        = New-Object System.Drawing.Font("Consolas",8.5)
$script:txtLogsView.Location    = New-Object System.Drawing.Point(24,100)
$script:txtLogsView.Size        = New-Object System.Drawing.Size(652,490)
$script:txtLogsView.WordWrap    = $false
$script:txtLogsView.BorderStyle = "None"
$panelLogs.Controls.Add($script:txtLogsView)

# ============================================================
# ABOUT PANEL
# ============================================================
$panelAbout = New-ContentPanel

$aboutTitle = New-Object System.Windows.Forms.Label
$aboutTitle.Text      = "About"
$aboutTitle.Font      = New-Object System.Drawing.Font("Segoe UI",16,[System.Drawing.FontStyle]::Bold)
$aboutTitle.ForeColor = $clrTextDark
$aboutTitle.Location  = New-Object System.Drawing.Point(24,20)
$aboutTitle.Size      = New-Object System.Drawing.Size(400,36)
$panelAbout.Controls.Add($aboutTitle)

$aboutCard = New-Object System.Windows.Forms.Panel
$aboutCard.BackColor = [System.Drawing.Color]::White
$aboutCard.Location  = New-Object System.Drawing.Point(24,68)
$aboutCard.Size      = New-Object System.Drawing.Size(652,300)
$panelAbout.Controls.Add($aboutCard)

function Add-AboutRow {
    param([string]$Heading,[string]$Value,[int]$Y)
    $h = New-Object System.Windows.Forms.Label
    $h.Text      = $Heading
    $h.Font      = New-Object System.Drawing.Font("Segoe UI",8.5)
    $h.ForeColor = $clrTextDim
    $h.Location  = New-Object System.Drawing.Point(20,$Y)
    $h.Size      = New-Object System.Drawing.Size(160,20)
    $aboutCard.Controls.Add($h)
    $v = New-Object System.Windows.Forms.Label
    $v.Text      = $Value
    $v.Font      = New-Object System.Drawing.Font("Segoe UI",8.5)
    $v.ForeColor = $clrTextDark
    $v.Location  = New-Object System.Drawing.Point(185,$Y)
    $v.Size      = New-Object System.Drawing.Size(440,20)
    $aboutCard.Controls.Add($v)
    return $v
}

$aboutAppName = New-Object System.Windows.Forms.Label
$aboutAppName.Text      = "EntryPass-MiHCM Sync"
$aboutAppName.Font      = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold)
$aboutAppName.ForeColor = $clrBlue
$aboutAppName.Location  = New-Object System.Drawing.Point(20,18)
$aboutAppName.Size      = New-Object System.Drawing.Size(400,30)
$aboutCard.Controls.Add($aboutAppName)

$aboutVer = New-Object System.Windows.Forms.Label
$aboutVer.Text      = "Version $script:appVersion"
$aboutVer.Font      = New-Object System.Drawing.Font("Segoe UI",9)
$aboutVer.ForeColor = $clrTextDim
$aboutVer.Location  = New-Object System.Drawing.Point(20,50)
$aboutVer.Size      = New-Object System.Drawing.Size(300,20)
$aboutCard.Controls.Add($aboutVer)

$aboutSep = New-Object System.Windows.Forms.Label
$aboutSep.BackColor = [System.Drawing.Color]::FromArgb(210,215,225)
$aboutSep.Location  = New-Object System.Drawing.Point(20,76)
$aboutSep.Size      = New-Object System.Drawing.Size(612,1)
$aboutCard.Controls.Add($aboutSep)

Add-AboutRow "Company"    "Dajayana Trading"       90  | Out-Null
Add-AboutRow "Website"    "www.dajayana.com"        112 | Out-Null
Add-AboutRow "Contact"    "+60 16-883 8338"         134 | Out-Null
$script:aboutLicRow = Add-AboutRow "License"  "Not validated"  156
Add-AboutRow "Config file" $script:configFile        178 | Out-Null
Add-AboutRow "Log file"    $script:logFile           200 | Out-Null
Add-AboutRow "Script dir"  $script:appDir            222 | Out-Null

$aboutDesc = New-Object System.Windows.Forms.Label
$aboutDesc.Text      = "Syncs EntryPass attendance data to MiHCM via REST API. Supports database direct mode, batch upload, duplicate detection, and automatic retry."
$aboutDesc.Font      = New-Object System.Drawing.Font("Segoe UI",8.5)
$aboutDesc.ForeColor = $clrTextDim
$aboutDesc.Location  = New-Object System.Drawing.Point(20,250)
$aboutDesc.Size      = New-Object System.Drawing.Size(610,44)
$aboutDesc.AutoSize  = $false
$aboutCard.Controls.Add($aboutDesc)

$btnCheckUpdate = New-Object System.Windows.Forms.Button
$btnCheckUpdate.Text      = "Check for Updates"
$btnCheckUpdate.Font      = New-Object System.Drawing.Font("Segoe UI",9)
$btnCheckUpdate.Location  = New-Object System.Drawing.Point(20,300)
$btnCheckUpdate.Size      = New-Object System.Drawing.Size(160,34)
$btnCheckUpdate.FlatStyle = "Flat"
$btnCheckUpdate.BackColor = $clrBlue
$btnCheckUpdate.ForeColor = [System.Drawing.Color]::White
$btnCheckUpdate.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnCheckUpdate.add_Click({ Check-ForUpdate })
$aboutCard.Controls.Add($btnCheckUpdate)

# ============================================================
# SYSTEM TRAY
# ============================================================
$script:trayIcon = New-Object System.Windows.Forms.NotifyIcon
$script:trayIcon.Text    = "EntryPass-MiHCM Sync"
$script:trayIcon.Visible = $true

# Use built-in app icon
try {
    $script:trayIcon.Icon = [System.Drawing.SystemIcons]::Application
} catch {}

$trayMenu = New-Object System.Windows.Forms.ContextMenuStrip
$trayMenuOpen    = $trayMenu.Items.Add("Open")
$trayMenuSyncNow = $trayMenu.Items.Add("Sync Now")
$trayMenuSep     = $trayMenu.Items.Add("-")
$trayMenuExit    = $trayMenu.Items.Add("Exit")

$trayMenuOpen.add_Click({
    $mainForm.Show()
    $mainForm.WindowState = "Normal"
    $mainForm.BringToFront()
})
$trayMenuSyncNow.add_Click({ Start-SyncBackground })
$trayMenuExit.add_Click({
    $script:reallyExit = $true
    $script:trayIcon.Visible = $false
    $script:trayIcon.Dispose()
    $mainForm.Close()
})
$script:trayIcon.ContextMenuStrip = $trayMenu
$script:trayIcon.add_DoubleClick({
    $mainForm.Show()
    $mainForm.WindowState = "Normal"
    $mainForm.BringToFront()
})

# ============================================================
# NAVIGATION LOGIC
# ============================================================
$script:contentPanels = @{
    navDashboard = $panelDashboard
    navConfig    = $panelConfig
    navLogs      = $panelLogs
    navAbout     = $panelAbout
}

function Switch-Panel {
    param([string]$NavKey)
    foreach ($k in $script:contentPanels.Keys) {
        $script:contentPanels[$k].Visible = ($k -eq $NavKey)
    }
    foreach ($k in $script:navItems.Keys) {
        $lbl = $script:navItems[$k]
        if ($k -eq $NavKey) {
            $lbl.BackColor = $script:clrSidebarAct
            $lbl.ForeColor = [System.Drawing.Color]::White
        } else {
            $lbl.BackColor = $script:clrSidebar
            $lbl.ForeColor = [System.Drawing.Color]::FromArgb(200,210,225)
        }
    }
    $script:activePanelName = $NavKey
    if ($NavKey -eq "navLogs") { Refresh-LogView }
    if ($NavKey -eq "navDashboard") { Refresh-Dashboard }
    if ($NavKey -eq "navAbout") { Refresh-About }
}

foreach ($k in $script:navItems.Keys) {
    $thisKey = $k.ToString()
    $script:navItems[$k].Tag = $thisKey
    $script:navItems[$k].add_Click({ Switch-Panel $this.Tag })
    $script:navItems[$k].add_MouseEnter({
        $src = $args[0]
        if ($src.BackColor -ne $script:clrSidebarAct) {
            $src.BackColor = $script:clrSidebarHov
        }
    })
    $script:navItems[$k].add_MouseLeave({
        $src = $args[0]
        if ($src.BackColor -ne $script:clrSidebarAct) {
            $src.BackColor = $script:clrSidebar
        }
    })
}

# ============================================================
# DASHBOARD REFRESH
# ============================================================
function Refresh-Dashboard {
    $cfg = Load-AppConfig
    if ($cfg -and -not [string]::IsNullOrWhiteSpace($cfg.licenseKey) -and -not [string]::IsNullOrWhiteSpace($cfg.primaryKey)) {
        $dataSource = if ($cfg.dataSource) { $cfg.dataSource } else { "file" }
        if ($dataSource -eq "database") {
            $script:lblStatus.Text      = "Database Direct Mode"
            $script:lblStatus.ForeColor = $script:clrBlue
        } else {
            $script:lblStatus.Text      = "File Mode"
            $script:lblStatus.ForeColor = $script:clrGreen
        }
    } else {
        $script:lblStatus.Text      = "Not Configured"
        $script:lblStatus.ForeColor = $script:clrGrey
    }
    if ($cfg) {
        $dataSource = if ($cfg.dataSource) { $cfg.dataSource } else { "file" }
        $script:lblSiteClient.Text   = if ($cfg.clientName)   { $cfg.clientName }   else { "--" }
        $script:lblSiteLocation.Text = if ($cfg.location)     { "$($cfg.location) ($($cfg.locationDesc))" } else { "--" }
        if ($dataSource -eq "database") {
            $script:lblSiteFolder.Text = if ($cfg.databasePath) { "DB: $($cfg.databasePath)" } else { "--" }
        } else {
            $script:lblSiteFolder.Text = if ($cfg.sourceFolder) { $cfg.sourceFolder } else { "--" }
        }
    }
    if ($script:lastStats.Time) {
        $script:lblLastSync.Text    = "Last sync: $($script:lastStats.Time)"
        $script:lblSaved.Text       = "$($script:lastStats.Saved)"
        $script:lblSkipped.Text     = "$($script:lastStats.Skipped)"
        $script:lblFailed.Text      = "$($script:lastStats.Failed)"
    }
}

# ============================================================
# CONFIG PANEL: LOAD VALUES FROM FILE
# ============================================================
function Load-ConfigToForm {
    $cfg = Load-AppConfig
    if (-not $cfg) { return }
    $script:txtCfgLicense.Text     = if ($cfg.licenseKey)    { $cfg.licenseKey }    else { "" }
    $script:txtCfgPrimary.Text     = if ($cfg.primaryKey)    { $cfg.primaryKey }    else { "" }
    $script:txtCfgSecret.Text      = if ($cfg.secretKey)     { $cfg.secretKey }     else { "" }
    $script:txtCfgClient.Text      = if ($cfg.clientName)    { $cfg.clientName }    else { "" }
    $script:txtCfgLocation.Text    = if ($cfg.location)      { $cfg.location }      else { "" }
    $script:txtCfgLocationDesc.Text= if ($cfg.locationDesc)  { $cfg.locationDesc }  else { "" }
    $script:txtCfgSourceFolder.Text= if ($cfg.sourceFolder)  { $cfg.sourceFolder }  else { "" }
    if ($cfg.apiEndpoint -match "uat") {
        $script:cmbCfgEndpoint.SelectedIndex = 1
    } else {
        $script:cmbCfgEndpoint.SelectedIndex = 0
    }
    if ($cfg.scheduleEnabled) { $script:chkSchedule.Checked = $true }
    $freqMap = @{ "15min"="Every 15 minutes"; "30min"="Every 30 minutes"; "1hour"="Hourly"; "2hour"="Every 2 hours" }
    if ($cfg.scheduleFrequency -and $freqMap.ContainsKey($cfg.scheduleFrequency)) {
        $script:cmbFrequency.SelectedItem = $freqMap[$cfg.scheduleFrequency]
    }
    # Database fields
    $dataSource = if ($cfg.dataSource) { $cfg.dataSource } else { "database" }
    if ($dataSource -eq "database") {
        $script:rdoDbMode.Checked   = $true
    } else {
        $script:rdoFileMode.Checked = $true
    }
    $script:txtCfgDbPath.Text  = if ($cfg.databasePath) { $cfg.databasePath } else { "" }
    $script:txtCfgFbUser.Text  = if ($cfg.firebird -and $cfg.firebird.user)     { $cfg.firebird.user }     else { "SYSDBA" }
    $script:txtCfgFbPass.Text  = if ($cfg.firebird -and $cfg.firebird.password) { $cfg.firebird.password } else { "masterkey" }
    $script:txtCfgFbLib.Text   = if ($cfg.firebird -and $cfg.firebird.clientLibrary) { $cfg.firebird.clientLibrary } else { "" }
    $syncDays = if ($cfg.syncDays -gt 0) { [int]$cfg.syncDays } else { 2 }
    $script:numSyncDays.Value  = [Math]::Max(1,[Math]::Min(30,$syncDays))
    Update-DataSourceUI
}

# ============================================================
# CONFIG PANEL: BUILD CONFIG FROM FORM
# ============================================================
function Get-ConfigFromForm {
    $endpointUrl = if ($script:cmbCfgEndpoint.SelectedIndex -eq 1) { "https://api.mihcm.com/uat" } else { "https://api.mihcm.com" }
    $freqVal     = switch ($script:cmbFrequency.SelectedItem) {
        "Every 15 minutes" { "15min" }
        "Hourly"           { "1hour" }
        "Every 2 hours"    { "2hour" }
        default            { "30min" }
    }
    $dataSource = if ($script:rdoDbMode.Checked) { "database" } else { "file" }
    return [ordered]@{
        licenseKey        = $script:txtCfgLicense.Text.Trim()
        primaryKey        = $script:txtCfgPrimary.Text.Trim()
        secretKey         = $script:txtCfgSecret.Text.Trim()
        apiEndpoint       = $endpointUrl
        clientName        = $script:txtCfgClient.Text.Trim()
        location          = ($script:txtCfgLocation.Text.Trim()).ToUpper()
        locationDesc      = $script:txtCfgLocationDesc.Text.Trim()
        sourceFolder      = $script:txtCfgSourceFolder.Text.Trim()
        dataSource        = $dataSource
        databasePath      = $script:txtCfgDbPath.Text.Trim()
        syncDays          = [int]$script:numSyncDays.Value
        firebird          = [ordered]@{
            user          = $script:txtCfgFbUser.Text.Trim()
            password      = $script:txtCfgFbPass.Text.Trim()
            clientLibrary = $script:txtCfgFbLib.Text.Trim()
        }
        batchSize         = 80
        scheduleEnabled   = $script:chkSchedule.Checked
        scheduleFrequency = $freqVal
    }
}

function Validate-CfgForm {
    $req = @(
        @{Field=$script:txtCfgLicense.Text; Name="License Key"},
        @{Field=$script:txtCfgPrimary.Text; Name="Primary Key"},
        @{Field=$script:txtCfgSecret.Text;  Name="Secret Key"},
        @{Field=$script:txtCfgClient.Text;  Name="Client Name"},
        @{Field=$script:txtCfgLocation.Text;Name="Location Code"}
    )
    foreach ($r in $req) {
        if ([string]::IsNullOrWhiteSpace($r.Field)) {
            [System.Windows.Forms.MessageBox]::Show("$($r.Name) is required.", "Validation", "OK", "Warning") | Out-Null
            return $false
        }
    }
    if ($script:rdoDbMode.Checked) {
        if ([string]::IsNullOrWhiteSpace($script:txtCfgDbPath.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Database Path is required when using Database Direct mode.", "Validation", "OK", "Warning") | Out-Null
            return $false
        }
    } else {
        if ([string]::IsNullOrWhiteSpace($script:txtCfgSourceFolder.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Source Folder is required when using File-based mode.", "Validation", "OK", "Warning") | Out-Null
            return $false
        }
    }
    return $true
}

# ============================================================
# LOGS VIEW REFRESH
# ============================================================
function Refresh-LogView {
    if (Test-Path $script:logFile) {
        try {
            $content = [System.IO.File]::ReadAllText($script:logFile, [System.Text.Encoding]::UTF8)
            $script:txtLogsView.Text = $content
            $script:txtLogsView.SelectionStart  = $script:txtLogsView.Text.Length
            $script:txtLogsView.ScrollToCaret()
        } catch {
            $script:txtLogsView.Text = "Cannot read log file: $_"
        }
    } else {
        $script:txtLogsView.Text = "No log file found. Run a sync to generate logs."
    }
}

# ============================================================
# ABOUT REFRESH
# ============================================================
function Refresh-About {
    $cfg = Load-AppConfig
    if ($cfg -and -not [string]::IsNullOrWhiteSpace($cfg.licenseKey)) {
        $script:aboutLicRow.Text = "Key: $($cfg.licenseKey.Substring(0,[Math]::Min(8,$cfg.licenseKey.Length)))..."
    } else {
        $script:aboutLicRow.Text = "Not configured"
    }
}

# ============================================================
# BACKGROUND SYNC
# ============================================================
function Start-SyncBackground {
    if (-not (Is-Configured)) {
        [System.Windows.Forms.MessageBox]::Show("Please complete the configuration first.", "Not Configured", "OK", "Warning") | Out-Null
        Switch-Panel "navConfig"
        return
    }
    if ($script:syncRunning) {
        [System.Windows.Forms.MessageBox]::Show("A sync is already running. Please wait.", "Busy", "OK", "Information") | Out-Null
        return
    }

    $script:syncRunning = $true
    $script:txtLiveLog.Text = ""
    $script:lblSyncStatus.Text      = "Syncing..."
    $script:lblSyncStatus.ForeColor = $clrOrange
    $script:btnSyncNow.Enabled      = $false
    $mainForm.Refresh()

    try {
        $res = Run-FullSync -Worker $null
        $now = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

        if ($res -and $res.Stats) {
            $script:lastStats.Saved   = $res.Stats.Saved
            $script:lastStats.Skipped = $res.Stats.Skipped
            $script:lastStats.Failed  = $res.Stats.Failed
            $script:lastStats.Time    = $now
            $script:lastStats.Result  = if ($res.Result) { $res.Result } else { "Done" }
            $script:lblSaved.Text   = "$($res.Stats.Saved)"
            $script:lblSkipped.Text = "$($res.Stats.Skipped)"
            $script:lblFailed.Text  = "$($res.Stats.Failed)"
            $script:lblLastSync.Text = "Last sync: $now"
            $msg = "Saved:$($res.Stats.Saved) Skipped:$($res.Stats.Skipped) Failed:$($res.Stats.Failed)"
            $script:lblSyncStatus.Text      = "Sync complete -- $msg"
            $script:lblSyncStatus.ForeColor = $clrGreen
            try { $script:trayIcon.ShowBalloonTip(5000,"EntryPass Sync","Sync complete: $msg",[System.Windows.Forms.ToolTipIcon]::Info) } catch {}
        } else {
            $script:lblSyncStatus.Text      = "Sync complete"
            $script:lblSyncStatus.ForeColor = $clrGreen
            try { $script:trayIcon.ShowBalloonTip(3000,"EntryPass Sync","Sync complete",[System.Windows.Forms.ToolTipIcon]::Info) } catch {}
        }
    } catch {
        $script:lblSyncStatus.Text      = "Sync error: $($_.Exception.Message)"
        $script:lblSyncStatus.ForeColor = [System.Drawing.Color]::FromArgb(210,60,60)
        try { $script:trayIcon.ShowBalloonTip(5000,"EntryPass Sync","Sync failed: $($_.Exception.Message)",[System.Windows.Forms.ToolTipIcon]::Error) } catch {}
    } finally {
        $script:syncRunning = $false
        $script:btnSyncNow.Enabled = $true
        Refresh-Dashboard
    }
}

# ============================================================
# WIRE UP BUTTON EVENTS
# ============================================================

# Dashboard: Sync Now
$script:btnSyncNow.add_Click({ Start-SyncBackground })

# Config: Browse folder (file mode)
$script:btnSyncFilesNow.add_Click({
    # Force a file-based sync using the Source Folder path from config
    $folder = $script:txtCfgSourceFolder.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path $folder)) {
        [System.Windows.Forms.MessageBox]::Show("Please set a valid Source Folder first.", "No Folder", "OK", "Warning") | Out-Null
        return
    }
    # Temporarily override dataSource to file, run sync, then restore
    $cfg = Load-AppConfig
    $origSource = $cfg.dataSource
    $cfg.dataSource = "file"
    $cfg.sourceFolder = $folder
    Save-AppConfig $cfg
    Switch-Panel "navDashboard"
    Start-SyncBackground
    # Restore original dataSource after sync starts (config reloaded inside sync)
    # No restore needed -- sync reads config at start, user can switch back manually
})

$script:btnBrowseFolder.add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description         = "Select folder containing EntryPass DATA*.txt files"
    $dlg.ShowNewFolderButton = $true
    if ($script:txtCfgSourceFolder.Text -and (Test-Path $script:txtCfgSourceFolder.Text)) {
        $dlg.SelectedPath = $script:txtCfgSourceFolder.Text
    }
    if ($dlg.ShowDialog() -eq "OK") {
        $script:txtCfgSourceFolder.Text = $dlg.SelectedPath
    }
})

# Config: Browse database path (database mode)
$script:btnBrowseDbPath.add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description         = "Select EntryPass P1_Server (or database root) folder containing the 'event' subfolder with TRANS.FDB"
    $dlg.ShowNewFolderButton = $false
    if ($script:txtCfgDbPath.Text -and (Test-Path $script:txtCfgDbPath.Text)) {
        $dlg.SelectedPath = $script:txtCfgDbPath.Text
    }
    if ($dlg.ShowDialog() -eq "OK") {
        $script:txtCfgDbPath.Text = $dlg.SelectedPath
    }
})

# Config: Test DB Connection
$script:btnTestDbConn.add_Click({
    $dbPath  = $script:txtCfgDbPath.Text.Trim()
    $fbUser  = $script:txtCfgFbUser.Text.Trim()
    $fbPass  = $script:txtCfgFbPass.Text.Trim()
    $fbLib   = $script:txtCfgFbLib.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($dbPath)) {
        $script:lblDbConnStatus.Text      = "Enter Database Path first."
        $script:lblDbConnStatus.ForeColor = $clrOrange
        return
    }
    $script:lblDbConnStatus.Text      = "Testing connection..."
    $script:lblDbConnStatus.ForeColor = $clrTextDim
    $panelConfig.Refresh()
    $result = Test-FirebirdConnection -DbPath $dbPath -FbUser $fbUser -FbPassword $fbPass -ClientLibrary $fbLib
    if ($result.Success) {
        $script:lblDbConnStatus.Text      = $result.Message
        $script:lblDbConnStatus.ForeColor = $clrGreen
    } else {
        $script:lblDbConnStatus.Text      = $result.Message
        $script:lblDbConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(210,60,60)
    }
})

# Config: Validate License
$script:btnValidateLicense.add_Click({
    $key = $script:txtCfgLicense.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($key)) {
        $script:lblCfgStatus.Text      = "Enter a License Key first."
        $script:lblCfgStatus.ForeColor = $clrOrange
        return
    }
    $script:lblCfgStatus.Text      = "Validating license..."
    $script:lblCfgStatus.ForeColor = $clrTextDim
    $panelConfig.Refresh()
    try {
        $r    = Invoke-WebRequest -Uri $script:licenseUrl -UseBasicParsing -TimeoutSec 10
        $data = $r.Content | ConvertFrom-Json
        $entry = $data | Where-Object { $_.licenseKey -eq $key }
        if (-not $entry) {
            $script:lblCfgStatus.Text      = "License NOT FOUND. Check key or contact Dajayana Trading."
            $script:lblCfgStatus.ForeColor = [System.Drawing.Color]::FromArgb(210,60,60)
            return
        }
        if ($entry.active -ne $true) {
            $script:lblCfgStatus.Text      = "License is INACTIVE. Contact Dajayana Trading."
            $script:lblCfgStatus.ForeColor = [System.Drawing.Color]::FromArgb(210,60,60)
            return
        }
        $exp = [datetime]::Parse($entry.expires)
        if ($exp -lt (Get-Date)) {
            $script:lblCfgStatus.Text      = "License EXPIRED ($($exp.ToString('yyyy-MM-dd'))). Contact Dajayana Trading."
            $script:lblCfgStatus.ForeColor = [System.Drawing.Color]::FromArgb(210,60,60)
            return
        }
        $script:lblCfgStatus.Text      = "License VALID. Client: $($entry.client). Expires: $($exp.ToString('yyyy-MM-dd'))."
        $script:lblCfgStatus.ForeColor = $clrGreen
    } catch {
        $script:lblCfgStatus.Text      = "Cannot reach license server: $($_.Exception.Message)"
        $script:lblCfgStatus.ForeColor = $clrOrange
    }
})

# Config: Test Connection
$script:btnTestConn.add_Click({
    $cfg = Load-AppConfig
    $script:lblCfgStatus.Text      = "Testing API connection..."
    $script:lblCfgStatus.ForeColor = $clrTextDim
    $panelConfig.Refresh()

    $primary = if ($cfg.primaryKey) { $cfg.primaryKey } else { "" }
    $secret  = if ($cfg.secretKey)  { $cfg.secretKey }  else { "" }
    $baseUrl = if ($cfg.apiEndpoint) { $cfg.apiEndpoint } else { "https://api.mihcm.com" }

    if ([string]::IsNullOrWhiteSpace($primary) -or [string]::IsNullOrWhiteSpace($secret)) {
        $script:lblCfgStatus.Text      = "API keys not configured in config.json"
        $script:lblCfgStatus.ForeColor = $clrOrange
        return
    }
    try {
        $url = "$baseUrl/oauth2/token?grantType=client_credentials&clientId=$primary&clientSecret=$secret"
        $raw = Invoke-WebRequest -Uri $url -Method GET -Headers @{"Ocp-Apim-Subscription-Key"=$primary} -UseBasicParsing -TimeoutSec 15
        $res = $raw.Content | ConvertFrom-Json
        if ($res.accessToken) {
            $script:lblCfgStatus.Text      = "API Connected -- Token obtained successfully"
            $script:lblCfgStatus.ForeColor = $clrGreen
        } else {
            $script:lblCfgStatus.Text      = "API responded but no token returned. Check keys."
            $script:lblCfgStatus.ForeColor = $clrOrange
        }
    } catch {
        $sc = ""
        if ($_.Exception.Response) { $sc = " (HTTP $([int]$_.Exception.Response.StatusCode))" }
        $script:lblCfgStatus.Text      = "API FAILED$sc -- $($_.Exception.Message)"
        $script:lblCfgStatus.ForeColor = [System.Drawing.Color]::FromArgb(210,60,60)
    }
})

# Config: Save Only
$script:btnSaveOnly.add_Click({
    if (-not (Validate-CfgForm)) { return }
    try {
        Save-AppConfig (Get-ConfigFromForm)
        $script:lblCfgStatus.Text      = "Configuration saved to: $script:configFile"
        $script:lblCfgStatus.ForeColor = $clrGreen
        Refresh-Dashboard
    } catch {
        $script:lblCfgStatus.Text      = "ERROR saving config: $_"
        $script:lblCfgStatus.ForeColor = [System.Drawing.Color]::FromArgb(210,60,60)
    }
})

# Config: Save & Install Schedule
$script:btnSaveInstall.add_Click({
    if (-not (Validate-CfgForm)) { return }
    try {
        $cfg = Get-ConfigFromForm
        Save-AppConfig $cfg
        $script:lblCfgStatus.Text      = "Config saved. Installing scheduled task..."
        $script:lblCfgStatus.ForeColor = $clrTextDim
        $panelConfig.Refresh()

        if ($cfg.scheduleEnabled) {
            # Task Scheduler with SYSTEM account requires admin -- re-launch elevated if needed
            $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
            if (-not $isAdmin) {
                $ans = [System.Windows.Forms.MessageBox]::Show("Installing a scheduled task requires Administrator privileges.`n`nClick Yes to restart as Administrator and install the task.","Admin Required","YesNo","Question")
                if ($ans -eq "Yes") {
                    Start-Process "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -NoProfile -File `"$script:appDir\EntryPassSync.ps1`"" -Verb RunAs
                    $script:reallyExit = $true
                    $mainForm.Close()
                }
                return
            }
            $result = Install-SyncTask -Location $cfg.location -Frequency $cfg.scheduleFrequency
            if ($result.Success) {
                $script:lblCfgStatus.Text      = "Saved and scheduled task installed: $($result.Message)"
                $script:lblCfgStatus.ForeColor = $clrGreen
                [System.Windows.Forms.MessageBox]::Show("Config saved and task installed.`n`n$($result.Message)","Setup Complete","OK","Information") | Out-Null
            } else {
                $script:lblCfgStatus.Text      = "Saved but task install failed. See details."
                $script:lblCfgStatus.ForeColor = $clrOrange
                [System.Windows.Forms.MessageBox]::Show("Config saved but task failed:`n$($result.Message)","Task Install Failed","OK","Warning") | Out-Null
            }
        } else {
            # Remove existing task if schedule was disabled
            Remove-SyncTask -Location $cfg.location | Out-Null
            $script:lblCfgStatus.Text      = "Config saved. Scheduled sync is disabled."
            $script:lblCfgStatus.ForeColor = $clrGreen
        }
        Refresh-Dashboard
    } catch {
        $script:lblCfgStatus.Text      = "ERROR: $_"
        $script:lblCfgStatus.ForeColor = [System.Drawing.Color]::FromArgb(210,60,60)
    }
})

# Logs: Refresh
$script:btnRefreshLog.add_Click({ Refresh-LogView })

# Logs: Clear
$script:btnClearLog.add_Click({
    $confirm = [System.Windows.Forms.MessageBox]::Show("Clear all log contents?","Confirm Clear","YesNo","Warning")
    if ($confirm -eq "Yes") {
        try {
            [System.IO.File]::WriteAllText($script:logFile, "", [System.Text.Encoding]::UTF8)
            $script:txtLogsView.Text = ""
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Cannot clear log: $_","Error","OK","Error") | Out-Null
        }
    }
})

# Logs: Open in Notepad
$script:btnOpenNotepad.add_Click({
    if (Test-Path $script:logFile) {
        Start-Process "notepad.exe" -ArgumentList "`"$script:logFile`""
    } else {
        [System.Windows.Forms.MessageBox]::Show("Log file not found. Run a sync first.","No Log","OK","Information") | Out-Null
    }
})

# ============================================================
# FORM EVENTS
# ============================================================
$mainForm.add_FormClosing({
    param($sender,$e)
    if (-not $script:reallyExit) {
        $e.Cancel = $true
        $mainForm.Hide()
        try { $script:trayIcon.ShowBalloonTip(2000,"EntryPass Sync","Running in the system tray. Double-click to restore.",[System.Windows.Forms.ToolTipIcon]::Info) } catch {}
    } else {
        $script:trayIcon.Visible = $false
        try { $script:trayIcon.Dispose() } catch {}
    }
})

$mainForm.add_Resize({
    if ($mainForm.WindowState -eq "Minimized") {
        $mainForm.Hide()
    }
})

$mainForm.add_Load({
    # Load config into config form
    Load-ConfigToForm

    # First launch: show config if not configured
    if (-not (Is-Configured)) {
        Switch-Panel "navConfig"
        $script:lblCfgStatus.Text      = "Please fill in all required fields to get started."
        $script:lblCfgStatus.ForeColor = $clrOrange
    } else {
        Switch-Panel "navDashboard"
    }
    Refresh-Dashboard
})

$mainForm.ResumeLayout($false)

# ============================================================
# AUTO-UPDATE CHECK
# ============================================================
function Check-ForUpdate {
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $wc = New-Object System.Net.WebClient
        $json = $wc.DownloadString($script:updateUrl)
        $remote = ConvertFrom-Json $json
        $remoteVer = [version]$remote.version
        $localVer  = [version]$script:appVersion
        if ($remoteVer -gt $localVer) {
            Write-SyncLog "Update available: v$($remote.version) (current: v$script:appVersion)"
            $ans = [System.Windows.Forms.MessageBox]::Show(
                "A new version is available!`n`nCurrent: v$script:appVersion`nNew: v$($remote.version)`n`n$($remote.notes)`n`nUpdate now?",
                "Update Available", "YesNo", "Information")
            if ($ans -eq "Yes") {
                Write-SyncLog "Downloading update v$($remote.version)..."
                $scriptPath = Join-Path $script:appDir "EntryPassSync.ps1"
                $tempPath   = Join-Path $script:appDir "EntryPassSync.ps1.update"
                $wc.DownloadFile($script:scriptUrl, $tempPath)
                # Verify download is valid (must contain our marker)
                $content = Get-Content $tempPath -Raw -ErrorAction Stop
                if ($content -match 'appVersion') {
                    Copy-Item $tempPath $scriptPath -Force
                    Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
                    Write-SyncLog "Update installed. Restarting..."
                    $psExe = "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
                    Start-Process $psExe -ArgumentList "-ExecutionPolicy Bypass -NoProfile -File `"$scriptPath`""
                    $script:reallyExit = $true
                    $mainForm.Close()
                } else {
                    Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
                    Write-SyncLog "Update download invalid — keeping current version"
                }
            }
        } else {
            Write-SyncLog "Version check: v$script:appVersion is up to date"
        }
    } catch {
        Write-SyncLog "Update check failed: $($_.Exception.Message)"
    }
}

# Check for updates shortly after startup (deferred so form is ready)
$script:updateTimer = New-Object System.Windows.Forms.Timer
$script:updateTimer.Interval = 3000  # 3 seconds after launch
$script:updateTimer.add_Tick({
    $script:updateTimer.Stop()
    Check-ForUpdate
})
$script:updateTimer.Start()

# ============================================================
# BUILT-IN SYNC TIMER (syncs every 15 minutes while app is open)
# ============================================================
$script:syncTimer = New-Object System.Windows.Forms.Timer
$script:syncTimer.Interval = 15 * 60 * 1000  # 15 minutes in ms
$script:syncTimer.add_Tick({
    if (-not $script:syncRunning -and (Is-Configured)) {
        Write-SyncLog "Auto-sync triggered (15-minute interval)"
        Start-SyncBackground
    }
})
$script:syncTimer.Start()

# Also add to Windows startup so it launches automatically
$startupKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run"
$startupName = "EntryPassMiHCMSync"
$startupCmd = "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -NoProfile -File `"$($script:appDir)\EntryPassSync.ps1`""
try {
    Set-ItemProperty -Path $startupKey -Name $startupName -Value $startupCmd -ErrorAction SilentlyContinue
} catch {}

# ============================================================
# RUN
# ============================================================
[System.Windows.Forms.Application]::Run($mainForm)

# Cleanup
try {
    $script:trayIcon.Visible = $false
    $script:trayIcon.Dispose()
} catch {}
