[CmdletBinding()]
param(
    [string]$CustomerName   = "Customer",
    [string]$ProjectName    = "Exchange Online Migration",
    [string]$ReportVersion  = "v1.0",

    [string]$OutputPath,

    # Optional – override auto-detection of latest *_OnPrem_*.csv in 4-Export / Export CSV
    [string]$OnPremRoomsCsv,
    [string]$OnPremRoomListsCsv,
    [string]$OnPremRoomListMembersCsv
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Paths & folders
# ---------------------------------------------------------------------------
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path      # ...\Rooms_Migration\1-Scripts
$baseFolder = Split-Path -Parent $scriptRoot                       # ...\Rooms_Migration

if (-not $OutputPath) {
    $OutputPath = Join-Path $baseFolder '2-Out'
}

$logFolder             = Join-Path $baseFolder '3-Logs'
$exportFolderPrimary   = Join-Path $baseFolder '4-Export'
$exportFolderSecondary = Join-Path (Split-Path $baseFolder -Parent) 'Export CSV'

foreach ($folder in @($OutputPath, $logFolder, $exportFolderPrimary)) {
    if (-not (Test-Path $folder)) {
        New-Item -Path $folder -ItemType Directory -Force | Out-Null
    }
}

$scriptName = [IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$logFile    = Join-Path $logFolder ("{0}_{1:yyyy-MM-dd}.log" -f $scriptName, (Get-Date))

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[{0}] {1}" -f $timestamp, $Message
    Write-Host $line
    Add-Content -Path $logFile -Value $line
}

Write-Log "----- Rooms & Room Lists migration report started -----"
Write-Log ("OutputPath                 = {0}" -f $OutputPath)
Write-Log ("OnPremRoomsCsv             = {0}" -f $OnPremRoomsCsv)
Write-Log ("OnPremRoomListsCsv         = {0}" -f $OnPremRoomListsCsv)
Write-Log ("OnPremRoomListMembersCsv   = {0}" -f $OnPremRoomListMembersCsv)

function Get-LatestCsvFile {
    param(
        [string]$Folder,
        [string]$Pattern
    )
    if (-not (Test-Path $Folder)) { return $null }

    Get-ChildItem -Path $Folder -Filter $Pattern -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

# ---------------------------------------------------------------------------
# Resolve on-prem CSV paths (auto-pick latest if not supplied)
# ---------------------------------------------------------------------------
if (-not $OnPremRoomsCsv) {
    $candidate = Get-LatestCsvFile -Folder $exportFolderPrimary   -Pattern 'Rooms_OnPrem_*.csv'
    if (-not $candidate) {
        $candidate = Get-LatestCsvFile -Folder $exportFolderSecondary -Pattern 'Rooms_OnPrem_*.csv'
    }
    if (-not $candidate) { throw "OnPremRoomsCsv not found. Run script 1 or pass -OnPremRoomsCsv." }

    $OnPremRoomsCsv = $candidate.FullName
    Write-Log ("Resolved OnPremRoomsCsv          = {0}" -f $OnPremRoomsCsv)
}

if (-not $OnPremRoomListsCsv) {
    $candidate = Get-LatestCsvFile -Folder $exportFolderPrimary   -Pattern 'RoomLists_OnPrem_*.csv'
    if (-not $candidate) {
        $candidate = Get-LatestCsvFile -Folder $exportFolderSecondary -Pattern 'RoomLists_OnPrem_*.csv'
    }
    if (-not $candidate) { throw "OnPremRoomListsCsv not found. Run script 1 or pass -OnPremRoomListsCsv." }

    $OnPremRoomListsCsv = $candidate.FullName
    Write-Log ("Resolved OnPremRoomListsCsv      = {0}" -f $OnPremRoomListsCsv)
}

if (-not $OnPremRoomListMembersCsv) {
    $candidate = Get-LatestCsvFile -Folder $exportFolderPrimary   -Pattern 'RoomListMembers_OnPrem_*.csv'
    if (-not $candidate) {
        $candidate = Get-LatestCsvFile -Folder $exportFolderSecondary -Pattern 'RoomListMembers_OnPrem_*.csv'
    }
    if (-not $candidate) { throw "OnPremRoomListMembersCsv not found. Run script 1 or pass -OnPremRoomListMembersCsv." }

    $OnPremRoomListMembersCsv = $candidate.FullName
    Write-Log ("Resolved OnPremRoomListMembersCsv= {0}" -f $OnPremRoomListMembersCsv)
}

$onPremExportFolder = Split-Path -Parent $OnPremRoomsCsv
Write-Log ("Using on-prem export folder: {0}" -f $onPremExportFolder)

# ---------------------------------------------------------------------------
# Connect to Exchange Online
# ---------------------------------------------------------------------------
try {
    Write-Log "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
    Write-Log "Connected to Exchange Online."
}
catch {
    Write-Log ("ERROR: Failed to connect to Exchange Online: {0}" -f $_.Exception.Message)
    throw
}

try {
    # -----------------------------------------------------------------------
    # Load on-prem data
    # -----------------------------------------------------------------------
    Write-Log "Loading on-prem data from CSVs..."
    $onPremRooms           = Import-Csv -Path $OnPremRoomsCsv
    $onPremRoomLists       = Import-Csv -Path $OnPremRoomListsCsv
    $onPremRoomListMembers = Import-Csv -Path $OnPremRoomListMembersCsv

    Write-Log ("On-prem rooms loaded: {0}" -f $onPremRooms.Count)
    Write-Log ("On-prem room lists loaded: {0}" -f $onPremRoomLists.Count)
    Write-Log ("On-prem room list members loaded: {0}" -f $onPremRoomListMembers.Count)

    # -----------------------------------------------------------------------
    # Load Exchange Online data
    # -----------------------------------------------------------------------
    Write-Log "Loading Exchange Online room mailboxes..."
    $exoRooms = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue |
                Select-Object DisplayName, PrimarySmtpAddress
    Write-Log ("EXO rooms loaded: {0}" -f ($exoRooms.Count))

    Write-Log "Loading Exchange Online room lists..."
    $exoRoomLists = Get-DistributionGroup -RecipientTypeDetails RoomList -ResultSize Unlimited -ErrorAction SilentlyContinue |
                    Select-Object DisplayName, PrimarySmtpAddress, Identity
    Write-Log ("EXO room lists loaded: {0}" -f ($exoRoomLists.Count))

    Write-Log "Loading EXO room list membership..."
    $exoRoomListMembers = @()
    foreach ($rl in $exoRoomLists) {
        Write-Log ("  Getting members for EXO room list: {0} [{1}]" -f $rl.DisplayName, $rl.Identity)
        $members = Get-DistributionGroupMember -Identity $rl.Identity -ResultSize Unlimited -ErrorAction SilentlyContinue |
                   Select-Object @{Name='RoomListDisplayName';        Expression = { $rl.DisplayName }},
                                 @{Name='RoomListPrimarySmtpAddress'; Expression = { $rl.PrimarySmtpAddress }},
                                 @{Name='MemberDisplayName';          Expression = { $_.Name }},
                                 @{Name='MemberPrimarySmtpAddress';   Expression = { $_.PrimarySmtpAddress }}
        if ($members) { $exoRoomListMembers += $members }
    }
    Write-Log ("EXO room list members loaded: {0}" -f ($exoRoomListMembers.Count))

    # -----------------------------------------------------------------------
    # Summary counts
    # -----------------------------------------------------------------------
    $onPremRoomCount      = $onPremRooms.Count
    $exoRoomCount         = $exoRooms.Count
    $onPremRoomListCount  = $onPremRoomLists.Count
    $exoRoomListCount     = $exoRoomLists.Count

    Write-Log "Summary counts:"
    Write-Log ("  On-prem rooms:      {0}" -f $onPremRoomCount)
    Write-Log ("  EXO rooms:          {0}" -f $exoRoomCount)
    Write-Log ("  On-prem room lists: {0}" -f $onPremRoomListCount)
    Write-Log ("  EXO room lists:     {0}" -f $exoRoomListCount)

    # -----------------------------------------------------------------------
    # Membership comparison (simple, based on known CSV columns)
    # -----------------------------------------------------------------------
    Write-Log "Comparing room list membership (on-prem vs EXO)..."

    $onPremRoomListMembersNorm = $onPremRoomListMembers | Where-Object {
        $_.RoomListPrimarySmtpAddress -and $_.MemberPrimarySmtpAddress
    } | ForEach-Object {
        [PSCustomObject]@{
            RoomListPrimarySmtpAddress = $_.RoomListPrimarySmtpAddress.ToLowerInvariant()
            MemberPrimarySmtpAddress   = $_.MemberPrimarySmtpAddress.ToLowerInvariant()
        }
    }

    $exoRoomListMembersNorm = $exoRoomListMembers | Where-Object {
        $_.RoomListPrimarySmtpAddress -and $_.MemberPrimarySmtpAddress
    } | ForEach-Object {
        [PSCustomObject]@{
            RoomListPrimarySmtpAddress = $_.RoomListPrimarySmtpAddress.ToString().ToLowerInvariant()
            MemberPrimarySmtpAddress   = $_.MemberPrimarySmtpAddress.ToString().ToLowerInvariant()
        }
    }

    $allListEmails = @(
        $onPremRoomListMembersNorm.RoomListPrimarySmtpAddress +
        $exoRoomListMembersNorm.RoomListPrimarySmtpAddress
    ) | Where-Object { $_ } | Sort-Object -Unique

    $mismatches = @()

    foreach ($listEmail in $allListEmails) {
        $onPremMembers = $onPremRoomListMembersNorm |
                         Where-Object { $_.RoomListPrimarySmtpAddress -eq $listEmail } |
                         Select-Object -ExpandProperty MemberPrimarySmtpAddress -Unique

        $exoMembers    = $exoRoomListMembersNorm  |
                         Where-Object { $_.RoomListPrimarySmtpAddress -eq $listEmail } |
                         Select-Object -ExpandProperty MemberPrimarySmtpAddress -Unique

        $missingInExo = @()
        $extraInExo   = @()

        if ($onPremMembers) {
            $missingInExo = $onPremMembers | Where-Object { $_ -notin $exoMembers }
        }
        if ($exoMembers) {
            $extraInExo   = $exoMembers    | Where-Object { $_ -notin $onPremMembers }
        }

        if ($missingInExo.Count -gt 0 -or $extraInExo.Count -gt 0) {
            $displayName =
                ($onPremRoomLists | Where-Object { $_.PrimarySmtpAddress -eq $listEmail } | Select-Object -First 1 -ExpandProperty DisplayName -ErrorAction SilentlyContinue) `
                -or
                ($exoRoomLists    | Where-Object { $_.PrimarySmtpAddress -eq $listEmail } | Select-Object -First 1 -ExpandProperty DisplayName -ErrorAction SilentlyContinue) `
                -or $listEmail

            $mismatches += [PSCustomObject]@{
                RoomListDisplayName        = $displayName
                RoomListPrimarySmtpAddress = $listEmail
                MissingInExo               = ($missingInExo -join '; ')
                ExtraInExo                 = ($extraInExo   -join '; ')
            }
        }
    }

    Write-Log ("Room lists with membership mismatches: {0}" -f $mismatches.Count)

    # -----------------------------------------------------------------------
    # Build HTML report (match sample layout)
    # -----------------------------------------------------------------------
    Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue | Out-Null

    $now                 = Get-Date
    $generatedDisplay    = $now.ToString('MM/dd/yyyy HH:mm:ss')
    $htmlPath            = Join-Path $OutputPath ("Rooms_RoomLists_Migration_Report_{0:yyyyMMdd_HHmm}.html" -f $now)

    Write-Log ("Generating HTML report: {0}" -f $htmlPath)

    $html = @()
    $html += '<!DOCTYPE html>'
    $html += '<html>'
    $html += '<head>'
    $html += '<meta charset="utf-8" />'
    $html += '<title>Rooms &amp; Room Lists Migration Report</title>'
    $html += '<style>'
    $html += ' body { font-family: Segoe UI, Arial, sans-serif; font-size: 11pt; margin: 20px; }'
    $html += ' h1, h2, h3 { color: #444444; }'
    $html += ' table { border-collapse: collapse; margin-top: 10px; margin-bottom: 20px; }'
    $html += ' th, td { border: 1px solid #cccccc; padding: 4px 8px; }'
    $html += ' th { background-color: #f0f0f0; }'
    $html += ' .kpi-table td:first-child { font-weight: 600; }'
    $html += ' .note { font-style: italic; color: #666666; }'
    $html += '</style>'
    $html += '</head>'
    $html += '<body>'

    $html += '<h1>Rooms &amp; Room Lists Migration Report</h1>'
    $html += ('<p><strong>Customer:</strong> {0}<br/>' -f [System.Web.HttpUtility]::HtmlEncode($CustomerName))
    $html += ('<strong>Project:</strong> {0}<br/>'     -f [System.Web.HttpUtility]::HtmlEncode($ProjectName))
    $html += ('<strong>Report version:</strong> {0}<br/>' -f [System.Web.HttpUtility]::HtmlEncode($ReportVersion))
    $html += ('<strong>Generated:</strong> {0}</p>' -f $generatedDisplay)

    # 1. Summary
    $html += '<h2>1. Summary</h2>'
    $html += '<table class="kpi-table">'
    $html += '  <tr><th>Metric</th><th>On-premises</th><th>Exchange Online</th></tr>'
    $html += ("  <tr><td>Total rooms</td><td>{0}</td><td>{1}</td></tr>" -f $onPremRoomCount, $exoRoomCount)
    $html += ("  <tr><td>Total room lists</td><td>{0}</td><td>{1}</td></tr>" -f $onPremRoomListCount, $exoRoomListCount)
    $html += '</table>'

    # 2. Details
    $html += '<h2>2. Details</h2>'
    $html += ("<p>Total rooms on-premises: <strong>{0}</strong>; total rooms in Exchange Online: <strong>{1}</strong>.<br/>" -f $onPremRoomCount, $exoRoomCount)
    $html += ("Total room lists on-premises: <strong>{0}</strong>; total room lists in Exchange Online: <strong>{1}</strong>.</p>" -f $onPremRoomListCount, $exoRoomListCount)
    $html += '<p class="note">The figures above are based on the latest on-premises CSV exports in '
    $html += '<code>4-Export</code> and live data from Exchange Online at the time this report was generated.</p>'

    # 2.1 Room mailboxes (detail) – table like sample
    $html += '<h3>2.1 Room mailboxes (detail)</h3>'
    $html += '<table>'
    $html += '  <tr><th>Room</th><th>On-prem primary SMTP</th><th>Present in EXO (synced or cloud)?</th><th>EXO primary SMTP</th></tr>'

    foreach ($room in $onPremRooms | Sort-Object DisplayName) {
        $roomName  = [System.Web.HttpUtility]::HtmlEncode($room.DisplayName)
        $roomSmtp  = [System.Web.HttpUtility]::HtmlEncode($room.PrimarySmtpAddress)
        $exoMatch  = $exoRooms | Where-Object {
            $_.PrimarySmtpAddress.ToString().ToLowerInvariant() -eq $room.PrimarySmtpAddress.ToString().ToLowerInvariant()
        } | Select-Object -First 1

        if ($exoMatch) {
            $existsInExo = "Yes"
            $exoSmtp     = [System.Web.HttpUtility]::HtmlEncode($exoMatch.PrimarySmtpAddress)
        }
        else {
            $existsInExo = "No"
            $exoSmtp     = ""
        }

        $html += ("  <tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>" -f $roomName, $roomSmtp, $existsInExo, $exoSmtp)
    }

    $html += '</table>'

    # 2.2 Room lists (detail)
    $html += '<h3>2.2 Room lists (detail)</h3>'
    $html += '<table>'
    $html += '  <tr><th>Room list</th><th>On-prem primary SMTP</th><th>Present in EXO (synced or cloud)?</th><th>EXO primary SMTP</th></tr>'

    foreach ($rl in $onPremRoomLists | Sort-Object DisplayName) {
        $name       = [System.Web.HttpUtility]::HtmlEncode($rl.DisplayName)
        $onPremSmtp = [System.Web.HttpUtility]::HtmlEncode($rl.PrimarySmtpAddress)

        $exoMatch   = $exoRoomLists | Where-Object {
            $_.PrimarySmtpAddress.ToString().ToLowerInvariant() -eq $rl.PrimarySmtpAddress.ToString().ToLowerInvariant()
        } | Select-Object -First 1

        if ($exoMatch) {
            $exists  = "Yes"
            $exoSmtp = [System.Web.HttpUtility]::HtmlEncode($exoMatch.PrimarySmtpAddress)
        }
        else {
            $exists  = "No"
            $exoSmtp = ""
        }

        $html += ("  <tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>" -f $name, $onPremSmtp, $exists, $exoSmtp)
    }

    $html += '</table>'

    # 2.3 Room lists with membership mismatches
    $html += '<h3>2.3 Room lists with membership mismatches</h3>'
    if ($mismatches.Count -eq 0) {
        # Match sample text exactly
        $html += '<p>No membership mismatches were detected. All on-premises room lists have matching '
        $html += 'membership (by primary SMTP address) in Exchange Online.</p>'
    }
    else {
        $html += '<table>'
        $html += '  <tr><th>Room list</th><th>On-prem primary SMTP</th><th>Missing in EXO (present on-prem)</th><th>Extra in EXO (not on-prem)</th></tr>'
        foreach ($m in $mismatches) {
            $name  = [System.Web.HttpUtility]::HtmlEncode($m.RoomListDisplayName)
            $email = [System.Web.HttpUtility]::HtmlEncode($m.RoomListPrimarySmtpAddress)
            $miss  = [System.Web.HttpUtility]::HtmlEncode($m.MissingInExo)
            $extra = [System.Web.HttpUtility]::HtmlEncode($m.ExtraInExo)

            $html += ("  <tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>" -f $name, $email, $miss, $extra)
        }
        $html += '</table>'
    }

    $html += '</body>'
    $html += '</html>'

    Set-Content -Path $htmlPath -Value ($html -join "`r`n") -Encoding UTF8
    Write-Log ("HTML report written to {0}" -f $htmlPath)
}
finally {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}
    Write-Log "Rooms & Room Lists migration report completed."
    Write-Log "----------------------------------------------------------"
}
