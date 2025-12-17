<#
.SYNOPSIS
Generate Rooms & Room Lists Migration HTML/DOCX report.

DEFAULTS:
  OnPrem*Csv  = latest *_OnPrem_*.csv in ..\..\Export CSV
  OutputPath  = ..\2-Out
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OnPremRoomsCsv,

    [Parameter(Mandatory = $false)]
    [string]$OnPremRoomListsCsv,

    [Parameter(Mandatory = $false)]
    [string]$OnPremRoomListMembersCsv,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$CustomerName = "Bank of Papua New Guinea",

    [Parameter(Mandatory = $false)]
    [string]$ProjectName = "Microsoft 365 Hybrid Messaging Upgrade",

    [Parameter(Mandatory = $false)]
    [string]$ReportTitle = "Rooms & Room Lists Migration Report",

    [Parameter(Mandatory = $false)]
    [string]$ReportVersion = "v1.0",

    [Parameter(Mandatory = $false)]
    [string]$ConsultantName = "Melky Warinak",

    [Parameter(Mandatory = $false)]
    [string]$ConsultantUrl = "https://myitsolutionspg.com",

    [Parameter(Mandatory = $false)]
    [string]$LogoUrl = "",   # optional: URL or local path

    [switch]$GenerateWord
)

#region Paths + logging setup
$scriptRoot         = Split-Path -Parent $MyInvocation.MyCommand.Path     # ...\Rooms_Migration\1-Scripts
$roomsMigrationRoot = Split-Path -Parent $scriptRoot                      # ...\Rooms_Migration
$projectRoot        = Split-Path -Parent $roomsMigrationRoot              # ...\ROOMS AND ROOM LIST MIGRATION
$exportCsvRoot      = Join-Path $projectRoot 'Export CSV'
$defaultOutPath     = Join-Path $roomsMigrationRoot '2-Out'
$logRoot            = Join-Path $roomsMigrationRoot '3-Logs'

if (-not (Test-Path -Path $logRoot)) {
    New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
}

$scriptBaseName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$logFile        = Join-Path $logRoot ("{0}_{1}.log" -f $scriptBaseName, (Get-Date -Format 'yyyy-MM-dd'))

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter()][ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line      = '{0} [{1}] {2}' -f $timestamp, $Level, $Message
    Add-Content -Path $logFile -Value $line
    Write-Host $Message
}
#endregion

#region Resolve paths
if (-not $OutputPath) { $OutputPath = $defaultOutPath }

if (-not (Test-Path -Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

if (-not $OnPremRoomsCsv) {
    $latest = Get-ChildItem -Path $exportCsvRoot -Filter 'Rooms_OnPrem_*.csv' -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latest) { $OnPremRoomsCsv = $latest.FullName }
}
if (-not $OnPremRoomListsCsv) {
    $latest = Get-ChildItem -Path $exportCsvRoot -Filter 'RoomLists_OnPrem_*.csv' -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latest) { $OnPremRoomListsCsv = $latest.FullName }
}
if (-not $OnPremRoomListMembersCsv) {
    $latest = Get-ChildItem -Path $exportCsvRoot -Filter 'RoomListMembers_OnPrem_*.csv' -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latest) { $OnPremRoomListMembersCsv = $latest.FullName }
}
#endregion

Write-Log "----- Rooms & Room Lists migration report started -----"
Write-Log "OutputPath = $OutputPath"
Write-Log "OnPremRoomsCsv = $OnPremRoomsCsv"
Write-Log "OnPremRoomListsCsv = $OnPremRoomListsCsv"
Write-Log "OnPremRoomListMembersCsv = $OnPremRoomListMembersCsv"

#region Output filenames
$timestamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$htmlPath  = Join-Path $OutputPath "Rooms_RoomLists_Migration_Report_$timestamp.html"
$docxPath  = Join-Path $OutputPath "Rooms_RoomLists_Migration_Report_$timestamp.docx"

Write-Host "Output folder: $OutputPath" -ForegroundColor Cyan
#endregion

#region EXO module & connect
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
    Write-Log "ExchangeOnlineManagement module not found." -Level ERROR
    throw "Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}

if (-not (Get-Module -Name ExchangeOnlineManagement)) {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
}

try {
    Get-ExoMailbox -ResultSize 1 -ErrorAction Stop | Out-Null
    Write-Host "Already connected to Exchange Online." -ForegroundColor Cyan
}
catch {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline
}
#endregion

#region Load on-prem CSVs
if (-not $OnPremRoomsCsv -or -not (Test-Path -Path $OnPremRoomsCsv)) {
    Write-Log "On-prem Rooms CSV not found. Expected Rooms_OnPrem_*.csv in '$exportCsvRoot' or pass -OnPremRoomsCsv." -Level ERROR
    throw "OnPremRoomsCsv not found."
}
if (-not $OnPremRoomListsCsv -or -not (Test-Path -Path $OnPremRoomListsCsv)) {
    Write-Log "On-prem Room Lists CSV not found. Expected RoomLists_OnPrem_*.csv in '$exportCsvRoot' or pass -OnPremRoomListsCsv." -Level ERROR
    throw "OnPremRoomListsCsv not found."
}
if (-not $OnPremRoomListMembersCsv -or -not (Test-Path -Path $OnPremRoomListMembersCsv)) {
    Write-Log "On-prem Room List Members CSV not found. Expected RoomListMembers_OnPrem_*.csv in '$exportCsvRoot' or pass -OnPremRoomListMembersCsv." -Level ERROR
    throw "OnPremRoomListMembersCsv not found."
}

$onPremRooms           = Import-Csv -Path $OnPremRoomsCsv
$onPremRoomLists       = Import-Csv -Path $OnPremRoomListsCsv
$onPremRoomListMembers = Import-Csv -Path $OnPremRoomListMembersCsv

Write-Host "Loaded on-prem CSVs:" -ForegroundColor Cyan
Write-Host "  Rooms:          $($onPremRooms.Count)"
Write-Host "  Room Lists:     $($onPremRoomLists.Count)"
Write-Host "  Room List Memb: $($onPremRoomListMembers.Count)"

Write-Log ("Loaded on-prem CSVs: Rooms={0}; RoomLists={1}; RoomListMembers={2}" -f `
           $onPremRooms.Count, $onPremRoomLists.Count, $onPremRoomListMembers.Count)
#endregion

#region Query EXO state
Write-Host "Querying EXO for room mailboxes..." -ForegroundColor Yellow
$exoRooms = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited |
    Select-Object DisplayName,PrimarySmtpAddress,Alias,Office,ResourceCapacity,RecipientTypeDetails,WhenCreated,WhenChanged

$exoRoomsBySmtp = @{}
foreach ($r in $exoRooms) {
    $smtp = $r.PrimarySmtpAddress.ToString().ToLower()
    if ($smtp -and -not $exoRoomsBySmtp.ContainsKey($smtp)) {
        $exoRoomsBySmtp[$smtp] = $r
    }
}

Write-Host "Querying EXO for room lists..." -ForegroundColor Yellow
$exoRoomLists = Get-DistributionGroup -RecipientTypeDetails RoomList -ResultSize Unlimited |
    Select-Object DisplayName,PrimarySmtpAddress,Alias,HiddenFromAddressListsEnabled,RequireSenderAuthenticationEnabled,WhenCreated,WhenChanged

$exoRoomListsBySmtp = @{}
$exoRoomListsByName = @{}
foreach ($rl in $exoRoomLists) {
    $smtp = $rl.PrimarySmtpAddress.ToString().ToLower()
    if ($smtp -and -not $exoRoomListsBySmtp.ContainsKey($smtp)) {
        $exoRoomListsBySmtp[$smtp] = $rl
    }
    $name = $rl.DisplayName
    if ($name -and -not $exoRoomListsByName.ContainsKey($name)) {
        $exoRoomListsByName[$name] = $rl
    }
}

Write-Log ("EXO state: Rooms={0}; RoomLists={1}" -f $exoRooms.Count, $exoRoomLists.Count)
#endregion

#region Room comparison
$roomReport = foreach ($room in $onPremRooms) {
    $onPremSmtp   = $room.PrimarySmtpAddress.ToString().ToLower()
    $onPremName   = $room.DisplayName
    $onPremOffice = $room.Office
    $onPremCap    = $room.ResourceCapacity

    $exoMatch = $null
    if ($onPremSmtp -and $exoRoomsBySmtp.ContainsKey($onPremSmtp)) {
        $exoMatch = $exoRoomsBySmtp[$onPremSmtp]
    }

    $status = if ($exoMatch) { "Migrated to EXO" } else { "Not migrated" }

    [PSCustomObject]@{
        OnPremDisplayName = $onPremName
        OnPremPrimarySmtp = $room.PrimarySmtpAddress
        OnPremOffice      = $onPremOffice
        OnPremCapacity    = $onPremCap
        ExoDisplayName    = if ($exoMatch) { $exoMatch.DisplayName } else { $null }
        ExoPrimarySmtp    = if ($exoMatch) { $exoMatch.PrimarySmtpAddress } else { $null }
        ExoOffice         = if ($exoMatch) { $exoMatch.Office } else { $null }
        ExoCapacity       = if ($exoMatch) { $exoMatch.ResourceCapacity } else { $null }
        Status            = $status
    }
}

$totalRooms       = $roomReport.Count
$roomsMigrated    = ($roomReport | Where-Object { $_.Status -eq "Migrated to EXO" }).Count
$roomsNotMigrated = $totalRooms - $roomsMigrated
#endregion

#region Room list comparison
$onPremListMemberCounts = $onPremRoomListMembers |
    Group-Object -Property RoomList |
    ForEach-Object {
        [PSCustomObject]@{
            RoomListSmtp = $_.Name
            MemberCount  = $_.Count
        }
    }

function Get-ExoRoomListMemberCount {
    param([string]$Identity)
    try {
        (Get-DistributionGroupMember -Identity $Identity -ResultSize Unlimited).Count
    }
    catch { 0 }
}

$roomListReport = foreach ($rl in $onPremRoomLists) {
    $onPremName = $rl.DisplayName
    $onPremSmtp = $rl.PrimarySmtpAddress.ToString().ToLower()

    $memberCountOnPrem = ($onPremListMemberCounts | Where-Object { $_.RoomListSmtp.ToLower() -eq $onPremSmtp }).MemberCount
    if (-not $memberCountOnPrem) { $memberCountOnPrem = 0 }

    $exo = $null
    if ($onPremSmtp -and $exoRoomListsBySmtp.ContainsKey($onPremSmtp)) {
        $exo = $exoRoomListsBySmtp[$onPremSmtp]
    }
    elseif ($exoRoomListsByName.ContainsKey($onPremName)) {
        $exo = $exoRoomListsByName[$onPremName]
    }

    $exoMemberCount = 0
    if ($exo) {
        $exoMemberCount = Get-ExoRoomListMemberCount -Identity $exo.PrimarySmtpAddress
    }

    $status = if (-not $exo) {
        "Not created in EXO"
    }
    elseif ($exoMemberCount -eq $memberCountOnPrem) {
        "Created in EXO (members match)"
    }
    else {
        "Created in EXO (membership mismatch)"
    }

    [PSCustomObject]@{
        OnPremDisplayName = $onPremName
        OnPremPrimarySmtp = $rl.PrimarySmtpAddress
        OnPremAlias       = $rl.Alias
        OnPremMembers     = $memberCountOnPrem
        ExoDisplayName    = if ($exo) { $exo.DisplayName } else { $null }
        ExoPrimarySmtp    = if ($exo) { $exo.PrimarySmtpAddress } else { $null }
        ExoAlias          = if ($exo) { $exo.Alias } else { $null }
        ExoMembers        = $exoMemberCount
        Status            = $status
    }
}

$totalRoomLists    = $roomListReport.Count
$roomListsInExo    = ($roomListReport | Where-Object { $_.Status -like "Created in EXO*" }).Count
$roomListsNotInExo = $totalRoomLists - $roomListsInExo
$roomListsMismatch = ($roomListReport | Where-Object { $_.Status -eq "Created in EXO (membership mismatch)" }).Count
#endregion

Write-Log ("Summary: RoomsMigrated={0}/{1}; RoomListsInEXO={2}/{3}; RoomListsMismatch={4}" -f `
           $roomsMigrated, $totalRooms, $roomListsInExo, $totalRoomLists, $roomListsMismatch)

#region Build HTML
Add-Type -AssemblyName System.Web
Write-Host "Building HTML report..." -ForegroundColor Yellow

$nowDisplay = Get-Date -Format "dd MMM yyyy HH:mm"

function Encode-Html {
    param([string]$Text)
    return [System.Web.HttpUtility]::HtmlEncode($Text)
}

$roomsRowsHtml = ($roomReport | Sort-Object Status,OnPremDisplayName | ForEach-Object {
    $statusClass = switch ($_.Status) {
        "Migrated to EXO" { "status-ok" }
        "Not migrated"    { "status-bad" }
        default           { "status-unknown" }
    }

@"
<tr>
  <td>$(Encode-Html $_.OnPremDisplayName)</td>
  <td>$(Encode-Html $_.OnPremPrimarySmtp)</td>
  <td>$(Encode-Html $_.OnPremOffice)</td>
  <td>$(Encode-Html $_.OnPremCapacity)</td>
  <td>$(Encode-Html $_.ExoDisplayName)</td>
  <td>$(Encode-Html $_.ExoPrimarySmtp)</td>
  <td class='$statusClass'>$(Encode-Html $_.Status)</td>
</tr>
"@
}) -join "`r`n"

$roomListsRowsHtml = ($roomListReport | Sort-Object Status,OnPremDisplayName | ForEach-Object {
    $statusClass = switch ($_.Status) {
        "Created in EXO (members match)"       { "status-ok" }
        "Created in EXO (membership mismatch)" { "status-warn" }
        "Not created in EXO"                   { "status-bad" }
        default                                { "status-unknown" }
    }

@"
<tr>
  <td>$(Encode-Html $_.OnPremDisplayName)</td>
  <td>$(Encode-Html $_.OnPremPrimarySmtp)</td>
  <td>$(Encode-Html $_.OnPremMembers)</td>
  <td>$(Encode-Html $_.ExoDisplayName)</td>
  <td>$(Encode-Html $_.ExoPrimarySmtp)</td>
  <td>$(Encode-Html $_.ExoMembers)</td>
  <td class='$statusClass'>$(Encode-Html $_.Status)</td>
</tr>
"@
}) -join "`r`n"

$logoHtml = if ($LogoUrl) {
    "<img src='$LogoUrl' alt='My IT Solutions Logo' class='logo' />"
} else {
    "<div class='logo-placeholder'>My IT Solutions</div>"
}

$html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>$ReportTitle - $CustomerName</title>
    <style>
        body {
            font-family: Segoe UI, Arial, sans-serif;
            margin: 0;
            padding: 0 0 40px 0;
            background-color: #f5f5f5;
        }
        .page {
            max-width: 1200px;
            margin: 0 auto;
            background-color: #ffffff;
            padding: 16px 24px 60px 24px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.1);
            position: relative;
            min-height: 95vh;
        }
        .header {
            display: grid;
            grid-template-columns: 1fr 2fr 1fr;
            align-items: center;
            margin-bottom: 12px;
        }
        .logo, .logo-placeholder {
            max-height: 48px;
        }
        .logo-placeholder {
            display: inline-block;
            padding: 8px 12px;
            border-radius: 6px;
            border: 1px solid #ccc;
            font-size: 12px;
            color: #666;
        }
        .header-center {
            text-align: center;
        }
        .header-center h1 {
            margin: 0;
            font-size: 20px;
        }
        .header-center .subtitle {
            font-size: 12px;
            color: #666;
        }
        .header-center .version {
            font-size: 11px;
            color: #999;
            margin-top: 2px;
        }
        .header-right {
            text-align: right;
            font-size: 11px;
            color: #555;
        }
        .header-right a {
            color: #0078d4;
            text-decoration: none;
        }
        .meta {
            font-size: 10px;
            color: #777;
            text-align: right;
            margin-top: 4px;
        }
        .kpi-row {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            grid-gap: 10px;
            margin: 18px 0;
        }
        .kpi {
            background: #fafafa;
            border-radius: 8px;
            padding: 10px 12px;
            border: 1px solid #e0e0e0;
        }
        .kpi-title {
            font-size: 11px;
            text-transform: uppercase;
            color: #777;
        }
        .kpi-value {
            font-size: 20px;
            font-weight: 600;
            margin-top: 4px;
        }
        .kpi-detail {
            font-size: 10px;
            color: #999;
            margin-top: 2px;
        }
        h2 {
            font-size: 14px;
            border-bottom: 1px solid #ddd;
            padding-bottom: 4px;
            margin-top: 18px;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin-top: 10px;
            font-size: 11px;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 4px 6px;
            text-align: left;
        }
        th {
            background-color: #f0f0f0;
            font-weight: 600;
        }
        tr:nth-child(even) {
            background-color: #fafafa;
        }
        .status-ok {
            color: #006400;
            font-weight: 600;
        }
        .status-bad {
            color: #b00020;
            font-weight: 600;
        }
        .status-warn {
            color: #e69100;
            font-weight: 600;
        }
        .status-unknown {
            color: #555;
        }
        .disclaimer-bar {
            position: fixed;
            bottom: 0;
            left: 0;
            right: 0;
            background: #333;
            color: #eee;
            font-size: 10px;
            padding: 5px 16px;
            text-align: center;
        }
        @media print {
            body {
                background: #ffffff;
            }
            .page {
                box-shadow: none;
                margin: 0;
            }
            .disclaimer-bar {
                position: static;
                margin-top: 8px;
            }
        }
    </style>
</head>
<body>
<div class="page">
    <div class="header">
        <div class="header-left">
            $logoHtml
        </div>
        <div class="header-center">
            <h1>$ReportTitle</h1>
            <div class="subtitle">$CustomerName – $ProjectName</div>
            <div class="version">$ReportVersion</div>
        </div>
        <div class="header-right">
            <div>By $ConsultantName</div>
            <div><a href="$ConsultantUrl">$ConsultantUrl</a></div>
        </div>
    </div>
    <div class="meta">
        Generated: $nowDisplay
    </div>

    <div class="kpi-row">
        <div class="kpi">
            <div class="kpi-title">Room Mailboxes (On-Prem vs EXO)</div>
            <div class="kpi-value">$roomsMigrated / $totalRooms</div>
            <div class="kpi-detail">$roomsNotMigrated remaining not yet migrated</div>
        </div>
        <div class="kpi">
            <div class="kpi-title">Room Lists (Created in EXO)</div>
            <div class="kpi-value">$roomListsInExo / $totalRoomLists</div>
            <div class="kpi-detail">$roomListsNotInExo on-prem lists not yet created in EXO</div>
        </div>
        <div class="kpi">
            <div class="kpi-title">EXO Room Lists Membership Mismatch</div>
            <div class="kpi-value">$roomListsMismatch</div>
            <div class="kpi-detail">Lists where EXO member count != on-prem member count</div>
        </div>
    </div>

    <h2>1. Rooms – On-Prem vs Exchange Online</h2>
    <p style="font-size:10px;color:#777;">
        Each row represents an on-prem room mailbox from the export. A room is considered "Migrated to EXO"
        when a matching room mailbox exists in Exchange Online with the same PrimarySmtpAddress.
    </p>
    <table>
        <thead>
            <tr>
                <th>On-Prem Display Name</th>
                <th>On-Prem Primary SMTP</th>
                <th>On-Prem Office</th>
                <th>On-Prem Capacity</th>
                <th>EXO Display Name</th>
                <th>EXO Primary SMTP</th>
                <th>Status</th>
            </tr>
        </thead>
        <tbody>
            $roomsRowsHtml
        </tbody>
    </table>

    <h2>2. Room Lists – On-Prem vs Exchange Online</h2>
    <p style="font-size:10px;color:#777;">
        Each row represents an on-prem Room List (distribution group with RecipientTypeDetails=RoomList).
        A list is considered CREATED IN EXO when a matching EXO Room List exists (by SMTP or DisplayName).
        Membership mismatch flags where the number of members differs between on-prem export and EXO.
    </p>
    <table>
        <thead>
            <tr>
                <th>On-Prem Room List</th>
                <th>On-Prem Primary SMTP</th>
                <th>On-Prem Members</th>
                <th>EXO Room List</th>
                <th>EXO Primary SMTP</th>
                <th>EXO Members</th>
                <th>Status</th>
            </tr>
        </thead>
        <tbody>
            $roomListsRowsHtml
        </tbody>
    </table>

</div>

<div class="disclaimer-bar">
    This report is auto-generated by My IT Solutions tooling for $CustomerName.
    Always validate results before performing production changes.
</div>
</body>
</html>
"@

Set-Content -Path $htmlPath -Value $html -Encoding UTF8
Write-Host "HTML report generated: $htmlPath" -ForegroundColor Green
Write-Log "HTML report generated at $htmlPath"
#endregion

#region Optional DOCX
if ($GenerateWord) {
    Write-Host "Attempting DOCX generation via Word COM..." -ForegroundColor Yellow
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $doc = $word.Documents.Open($htmlPath)
        $wdFormatXMLDocument = 12
        $doc.SaveAs([ref]$docxPath, [ref]$wdFormatXMLDocument)
        $doc.Close()
        $word.Quit()
        Write-Host "DOCX report generated: $docxPath" -ForegroundColor Green
        Write-Log "DOCX report generated at $docxPath"
    }
    catch {
        Write-Warning "DOCX generation failed: $($_.Exception.Message)"
        Write-Log "DOCX generation failed: $($_.Exception.Message)" -Level ERROR
        try { if ($word) { $word.Quit() } } catch {}
    }
}
#endregion

Write-Host "Rooms & Room Lists migration report complete." -ForegroundColor Cyan
Write-Log "Rooms & Room Lists migration report completed."
Write-Log "------------------------------------------------"
