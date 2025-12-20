<#
.SYNOPSIS
  Generates an HTML Rooms & Room Lists migration report (Exchange 2016 -> Exchange Online).

.DESCRIPTION
  - Loads on-prem export CSVs produced by 1-Export-RoomsAndRoomLists_OnPrem.ps1
  - Connects to Exchange Online and retrieves current Room Mailboxes + Room Lists + membership
  - Produces a CAB-friendly HTML report with summary, inventories, and per-roomlist membership detail

.NOTES
  Author: My IT Solutions (genericized for GitHub)
  Tested: Windows PowerShell 5.1 with ExchangeOnlineManagement module
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$CustomerName,

    [Parameter(Mandatory)]
    [string]$ProjectName,

    [Parameter(Mandatory)]
    [string]$ReportVersion,

    # Optional: if you want to explicitly point to CSVs (otherwise latest in 4-Export is auto-selected)
    [string]$OnPremRoomsCsv,
    [string]$OnPremRoomListsCsv,
    [string]$OnPremRoomListMembersCsv,

    # Optional: override output root (otherwise auto-resolved from script location)
    [string]$ProjectRoot
)

Set-StrictMode -Version Latest

$scriptVersion = "v1.3"
$ErrorActionPreference = 'Stop'

# Load System.Web for HTML encoding (Windows PowerShell)
try { Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue } catch {}

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR')] [string]$Level = 'INFO')
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[{0}] {1}: {2}" -f $ts, $Level, $Message
    Write-Host $line
    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $line
    }
}

function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Get-LatestFile {
    param(
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$Filter
    )
    if (-not (Test-Path -LiteralPath $Folder)) { return $null }
    $f = Get-ChildItem -LiteralPath $Folder -Filter $Filter -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    return $f.FullName
}

function Html-Enc([object]$Value) {
    if ($null -eq $Value) { return '' }
    return [System.Web.HttpUtility]::HtmlEncode([string]$Value)
}

# ----------------------------
# Resolve folder structure
# ----------------------------
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    # Script is expected to be in: <ProjectRoot>\1-Scripts
    $ProjectRoot = Split-Path -Parent $scriptDir
}

$OutPath   = Join-Path $ProjectRoot '2-Out'
$LogsPath  = Join-Path $ProjectRoot '3-Logs'
$ExportPath= Join-Path $ProjectRoot '4-Export'

Ensure-Folder -Path $OutPath
Ensure-Folder -Path $LogsPath

$runStamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
$script:LogFile = Join-Path $LogsPath ("4-New-RoomsAndRoomLists_MigrationReport_{0}.log" -f $runStamp)

Write-Log "----- Rooms & Room Lists migration report started -----"
Write-Log ("ProjectRoot = {0}" -f $ProjectRoot)
Write-Log ("ExportPath  = {0}" -f $ExportPath)
Write-Log ("OutPath     = {0}" -f $OutPath)
Write-Log ("LogsPath    = {0}" -f $LogsPath)

# ----------------------------
# Resolve CSV inputs
# ----------------------------
if ([string]::IsNullOrWhiteSpace($OnPremRoomsCsv)) {
    $OnPremRoomsCsv = Get-LatestFile -Folder $ExportPath -Filter 'Rooms_OnPrem_*.csv'
}
if ([string]::IsNullOrWhiteSpace($OnPremRoomListsCsv)) {
    $OnPremRoomListsCsv = Get-LatestFile -Folder $ExportPath -Filter 'RoomLists_OnPrem_*.csv'
}
if ([string]::IsNullOrWhiteSpace($OnPremRoomListMembersCsv)) {
    $OnPremRoomListMembersCsv = Get-LatestFile -Folder $ExportPath -Filter 'RoomListMembers_OnPrem_*.csv'
}

Write-Log ("Resolved OnPremRoomsCsv          = {0}" -f $OnPremRoomsCsv)
Write-Log ("Resolved OnPremRoomListsCsv      = {0}" -f $OnPremRoomListsCsv)
Write-Log ("Resolved OnPremRoomListMembersCsv= {0}" -f $OnPremRoomListMembersCsv)

foreach ($p in @($OnPremRoomsCsv,$OnPremRoomListsCsv,$OnPremRoomListMembersCsv)) {
    if ([string]::IsNullOrWhiteSpace($p) -or -not (Test-Path -LiteralPath $p)) {
        throw "Required CSV not found. Ensure you ran 1-Export-RoomsAndRoomLists_OnPrem.ps1 and that exports exist under: $ExportPath"
    }
}

# ----------------------------
# Load on-prem exports
# ----------------------------
Write-Log "Loading on-prem data from CSVs..."
$onPremRooms = Import-Csv -LiteralPath $OnPremRoomsCsv
$onPremRoomLists = Import-Csv -LiteralPath $OnPremRoomListsCsv
$onPremRoomListMembers = Import-Csv -LiteralPath $OnPremRoomListMembersCsv

# Normalize
$onPremRooms = $onPremRooms | Where-Object { $_.PrimarySmtpAddress } |
    Select-Object DisplayName, PrimarySmtpAddress, Alias, RecipientTypeDetails, Identity |
    Sort-Object DisplayName

$onPremRoomLists = $onPremRoomLists | Where-Object { $_.PrimarySmtpAddress } |
    Select-Object DisplayName, PrimarySmtpAddress, Alias, ManagedBy, RecipientTypeDetails |
    Sort-Object DisplayName

$onPremRoomListMembers = $onPremRoomListMembers | Where-Object { $_.RoomListPrimarySmtpAddress -and $_.MemberPrimarySmtpAddress } |
    Select-Object RoomListDisplayName, RoomListPrimarySmtpAddress, MemberDisplayName, MemberPrimarySmtpAddress, MemberRecipientTypeDetails |
    Sort-Object RoomListDisplayName, MemberDisplayName

Write-Log ("On-prem rooms loaded: {0}" -f (@($onPremRooms).Count))
Write-Log ("On-prem room lists loaded: {0}" -f (@($onPremRoomLists).Count))
Write-Log ("On-prem room list members loaded: {0}" -f (@($onPremRoomListMembers).Count))

# ----------------------------
# Connect to EXO and load cloud data
# ----------------------------
Write-Log "Connecting to Exchange Online..."
try {
    if (-not (Get-Module ExchangeOnlineManagement -ListAvailable)) {
        throw "ExchangeOnlineManagement module is not installed. Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Connect-ExchangeOnline -ShowBanner:$false | Out-Null
}
catch {
    throw "Failed to connect to Exchange Online. $($_.Exception.Message)"
}
Write-Log "Connected to Exchange Online."

Write-Log "Loading Exchange Online room mailboxes..."
$exoRooms = @()
try {
    if (Get-Command Get-EXOMailbox -ErrorAction SilentlyContinue) {
        $exoRooms = @(Get-EXOMailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited |
            Select-Object DisplayName, PrimarySmtpAddress, Alias, RecipientTypeDetails, Identity |
            Sort-Object DisplayName)
    } else {
        # Fallback for older cmdlet sets
        $exoRooms = @(Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited |
            Select-Object DisplayName, PrimarySmtpAddress, Alias, RecipientTypeDetails, Identity |
            Sort-Object DisplayName)
    }
} catch {
    Write-Log ("WARN: Failed to load EXO rooms. {0}" -f $_.Exception.Message) 'WARN'
}
Write-Log ("EXO rooms loaded: {0}" -f (@($exoRooms).Count))

Write-Log "Loading Exchange Online room lists..."
$exoRoomLists = @()
try {
    $exoRoomLists = @(Get-DistributionGroup -RecipientTypeDetails RoomList -ResultSize Unlimited -ErrorAction Stop |
        Select-Object DisplayName, PrimarySmtpAddress, Alias, RecipientTypeDetails, Identity |
        Sort-Object DisplayName)
} catch {
    Write-Log ("WARN: Failed to load EXO room lists. {0}" -f $_.Exception.Message) 'WARN'
}
Write-Log ("EXO room lists loaded: {0}" -f (@($exoRoomLists).Count))

Write-Log "Loading EXO room list membership..."
$exoRoomListMembers = @()
if (@($exoRoomLists).Count -gt 0) {
    foreach ($rl in $exoRoomLists) {
        try {
            Write-Log ("  Getting members for EXO room list: {0} [{1}]" -f $rl.DisplayName, $rl.Identity)
            $members = Get-DistributionGroupMember -Identity $rl.Identity -ResultSize Unlimited -ErrorAction Stop
            foreach ($m in $members) {
                $smtp = $null
                if ($m.PSObject.Properties.Match('PrimarySmtpAddress').Count -gt 0 -and $m.PrimarySmtpAddress) {
                    $smtp = [string]$m.PrimarySmtpAddress
                } elseif ($m.PSObject.Properties.Match('WindowsEmailAddress').Count -gt 0 -and $m.WindowsEmailAddress) {
                    $smtp = [string]$m.WindowsEmailAddress
                } elseif ($m.PSObject.Properties.Match('ExternalEmailAddress').Count -gt 0 -and $m.ExternalEmailAddress) {
                    $smtp = [string]$m.ExternalEmailAddress
                }

                $exoRoomListMembers += [pscustomobject]@{
                    RoomListDisplayName        = $rl.DisplayName
                    RoomListPrimarySmtpAddress = [string]$rl.PrimarySmtpAddress
                    MemberDisplayName          = $m.DisplayName
                    MemberPrimarySmtpAddress   = $smtp
                    MemberRecipientTypeDetails = $m.RecipientTypeDetails
                }
            }
        } catch {
            Write-Log ("WARN: Failed to get members for '{0}'. {1}" -f $rl.DisplayName, $_.Exception.Message) 'WARN'
        }
    }
}
$exoRoomListMembers = $exoRoomListMembers | Where-Object { $_.RoomListPrimarySmtpAddress -and $_.MemberPrimarySmtpAddress } |
    Sort-Object RoomListDisplayName, MemberDisplayName
Write-Log ("EXO room list members loaded: {0}" -f (@($exoRoomListMembers).Count))

# ----------------------------
# Compare membership mismatch (SMTP sets)
# ----------------------------
Write-Log "Comparing room list membership (on-prem vs EXO)..."

function Get-SmtpSet([object[]]$members) {
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($x in $members) {
        if (-not $x) { continue }
        $smtp = $x.MemberPrimarySmtpAddress
        if ([string]::IsNullOrWhiteSpace($smtp)) { continue }
        $null = $set.Add($smtp.ToLower())
    }
    return $set
}

$membershipMismatches = @()
foreach ($opl in $onPremRoomLists) {
    $k = ([string]$opl.PrimarySmtpAddress).ToLower()
    $opMembers = @($onPremRoomListMembers | Where-Object { ([string]$_.RoomListPrimarySmtpAddress).ToLower() -eq $k })
    $xoMembers = @($exoRoomListMembers | Where-Object { ([string]$_.RoomListPrimarySmtpAddress).ToLower() -eq $k })

    $opSet = Get-SmtpSet -members $opMembers
    $xoSet = Get-SmtpSet -members $xoMembers

    $onlyOnPrem = @()
    $onlyExo = @()

    foreach ($s in $opSet) { if (-not $xoSet.Contains($s)) { $onlyOnPrem += $s } }
    foreach ($s in $xoSet) { if (-not $opSet.Contains($s)) { $onlyExo += $s } }

    if ($onlyOnPrem.Count -gt 0 -or $onlyExo.Count -gt 0) {
        $membershipMismatches += [pscustomobject]@{
            RoomListDisplayName        = $opl.DisplayName
            RoomListPrimarySmtpAddress = $opl.PrimarySmtpAddress
            OnlyOnPremCount            = $onlyOnPrem.Count
            OnlyExoCount               = $onlyExo.Count
            OnlyOnPrem                 = ($onlyOnPrem -join ', ')
            OnlyExo                    = ($onlyExo -join ', ')
        }
    }
}

Write-Log ("Room lists with membership mismatches: {0}" -f ($membershipMismatches.Count))

# ----------------------------
# Build HTML report
# ----------------------------
$generatedOn = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
$outFile = Join-Path $OutPath ("Rooms_RoomLists_Migration_Report_{0}.html" -f $runStamp)

Write-Log ("Generating HTML report: {0}" -f $outFile)

$html = @()
$html += '<!DOCTYPE html>'
$html += '<html lang="en">'
$html += '<head>'
$html += '<meta charset="UTF-8" />'
$html += '<meta name="viewport" content="width=device-width, initial-scale=1.0" />'
$html += ("<title>{0} - Rooms & Room Lists Migration Report</title>" -f (Html-Enc $CustomerName))

# --- Styles ---
$html += '<style>'
$html += ':root{--bg:#f6f7fb;--card:#ffffff;--text:#0f172a;--muted:#000000;--border:#e5e7eb;--primary:#0019db;--accent:#2563eb;--accent2:#0f172a;}'
$html += 'body{margin:0;font-family:Segoe UI,Roboto,Arial,sans-serif;background:var(--bg);color:var(--text);}'
$html += '.wrap{max-width:1120px;margin:22px auto;padding:0 14px 92px;}'
$html += '.header{background:linear-gradient(135deg,#38bdf8 0%,#0ea5e9 55%,#0284c7 100%);color:#fff;padding:22px 18px 16px;border-bottom:1px solid rgba(255,255,255,.22);}'
$html += '.header-top{display:flex;gap:18px;justify-content:space-between;align-items:flex-end;flex-wrap:wrap;}'
$html += '.title{font-size:22px;font-weight:700;letter-spacing:.3px;margin:0;}'
$html += '.subtitle{margin:6px 0 0 0;font-size:14px;opacity:.88;line-height:1.4;}'
$html += '.stamp{font-size:14px;opacity:.88;text-align:right;line-height:1.4;}'
$html += '.nav{margin-top:14px;display:flex;gap:10px;flex-wrap:wrap;}'
$html += '.nav a{color:#100f5e;text-decoration:none;font-size:14px;font-weight: bolder;padding:6px 10px;border:1px solid rgba(27, 30, 32, 0.07);border-radius:999px;}'
$html += '.nav a:hover{background:rgba(37,99,235,.18);}'
$html += '.card{background:var(--card);border:1px solid var(--border);border-radius:14px;padding:14px 16px;margin-top:16px;box-shadow:0 6px 18px rgba(2,6,23,.06);}'
$html += '#summary.card{background:linear-gradient(180deg,#eff6ff 0%,#ffffff 55%);border-left:6px solid #38bdf8;}'
$html += '#inventory.card{background:linear-gradient(180deg,#f0fdf4 0%,#ffffff 55%);border-left:6px solid #22c55e;}'
$html += '#memberships.card{background:linear-gradient(180deg,#faf5ff 0%,#ffffff 55%);border-left:6px solid #a855f7;}'
$html += '#mismatches.card{background:linear-gradient(180deg,#fff7ed 0%,#ffffff 55%);border-left:6px solid #fb923c;}'
$html += '.card h2{margin:0 0 12px;font-size:15px;display:flex;align-items:center;gap:10px;}'
$html += '.card h2:before{content:"";display:inline-block;width:10px;height:10px;border-radius:999px;background:var(--primary);box-shadow:0 0 0 4px rgba(14,165,233,.18);}'
$html += 'h2{font-size:16px;margin:0 0 10px 0;}'
$html += 'h3{font-size:14px;margin:18px 0 10px 0;}'
$html += 'p{margin:8px 0;color:var(--muted);font-size:13px;line-height:1.5;}'
$html += '.grid{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;}'
$html += '@media(max-width:980px){.grid{grid-template-columns:repeat(2,1fr);}}'
$html += '.kpi{border:1px solid var(--border);border-radius:12px;padding:10px 12px;background:#fbfdff;}'
$html += '.kpi .label{font-size:14px;color:var(--muted);} .kpi .value{font-size:20px;font-weight:700;margin-top:2px;}'
$html += 'table{width:100%;border-collapse:collapse;margin-top:10px;font-size:14px;}'
$html += 'th,td{border:1px solid var(--border);padding:8px 10px;vertical-align:top;}'
$html += 'th{background:#f1f5f9;text-align:left;font-weight:700;}'
$html += '.small{font-size:14px;color:var(--muted);} .mono{font-family:ui-monospace,SFMono-Regular,Consolas,Monaco,monospace;}'
$html += '.pill{display:inline-block;font-size:14px;padding:2px 8px;border-radius:999px;border:1px solid var(--border);background:#3f8d53;color:var(--muted);}'
$html += '.two-col{display:grid;grid-template-columns:1fr 1fr;gap:12px;}'
$html += '@media(max-width:980px){.two-col{grid-template-columns:1fr;}}'
$html += '.footer-fixed{position:fixed;left:0;right:0;bottom:0;background:rgba(113, 124, 208, 0.92);backdrop-filter:saturate(180%) blur(10px);border-top:1px solid var(--border);box-shadow:0 -6px 18px rgba(2,6,23,.10);padding:10px 14px;z-index:50;}'
$html += '.footer-inner{max-width:1120px;margin:0 auto;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;font-size:14px;color:var(--muted);}'
$html += '.footer-inner a{color:var(--primary);text-decoration:none;font-weight:700;}'
$html += '.footer-inner a:hover{text-decoration:underline;}'
$html += '.footer-sep{opacity:.7;}'
$html += '</style>'
$html += '</head>'
$html += '<body>'
$html += '<div class="wrap">'

# --- Header ---
$html += '<div class="header">'
$html += '<div class="header-top">'
$html += ('<div><div class="title">Rooms &amp; Room Lists Migration Report</div><div class="subtitle"><span class="pill">{0}</span> &nbsp; <span class="pill">{1}</span> &nbsp; <span class="pill">{2}</span></div></div>' -f (Html-Enc $CustomerName), (Html-Enc $ProjectName), (Html-Enc $ReportVersion))
$html += ('<div class="stamp"><div><b>Generated:</b> {0}</div><div class="small">Source: Exchange 2016 exports + Exchange Online live queries</div></div>' -f (Html-Enc $generatedOn))
$html += '</div>'
$html += '<div class="nav">'
$html += '<a href="#summary">Summary</a>'
$html += '<a href="#inventory">Inventory</a>'
$html += '<a href="#memberships">Room list memberships</a>'
$html += '<a href="#mismatches">Mismatches</a>'
$html += '</div>'
$html += '</div>'

# --- Summary KPIs ---
$html += '<div class="card" id="summary">'
$html += '<h2>1. Summary</h2>'
$html += '<div class="grid">'
$html += ('<div class="kpi"><div class="label">On-prem rooms</div><div class="value">{0}</div></div>' -f @($onPremRooms).Count)
$html += ('<div class="kpi"><div class="label">EXO rooms</div><div class="value">{0}</div></div>' -f @($exoRooms).Count)
$html += ('<div class="kpi"><div class="label">On-prem room lists</div><div class="value">{0}</div></div>' -f @($onPremRoomLists).Count)
$html += ('<div class="kpi"><div class="label">EXO room lists</div><div class="value">{0}</div></div>' -f @($exoRoomLists).Count)
$html += ('<div class="kpi"><div class="label">Membership mismatches</div><div class="value">{0}</div></div>' -f $membershipMismatches.Count)
$html += '</div>'
$html += '<p class="small">If EXO rooms show <b>0</b> but on-prem rooms exist, confirm rooms have been migrated to Exchange Online (room mailboxes are mailboxes, not distribution groups).</p>'
$html += '</div>'

# --- Inventory section (names) ---
$html += '<div class="card" id="inventory">'
$html += '<h2>2. Inventory</h2>'

# 2.1 On-prem rooms
$html += '<h3>2.1 On-prem rooms</h3>'
$html += ('<p class="small">Source: <span class="mono">{0}</span></p>' -f (Html-Enc (Split-Path -Leaf $OnPremRoomsCsv)))
$html += '<table><thead><tr><th>Display Name</th><th>Primary SMTP</th><th>Alias</th></tr></thead><tbody>'
foreach ($r in $onPremRooms) {
    $html += ('<tr><td>{0}</td><td class="mono">{1}</td><td class="mono">{2}</td></tr>' -f (Html-Enc $r.DisplayName), (Html-Enc $r.PrimarySmtpAddress), (Html-Enc $r.Alias))
}
$html += '</tbody></table>'

# 2.2 EXO rooms
$html += '<h3>2.2 Exchange Online rooms</h3>'
if (@($exoRooms).Count -eq 0) {
    $html += '<p><span class="pill">INFO</span> No room mailboxes were returned from Exchange Online at runtime.</p>'
} else {
    $html += '<table><thead><tr><th>Display Name</th><th>Primary SMTP</th><th>Alias</th></tr></thead><tbody>'
    foreach ($r in $exoRooms) {
        $html += ('<tr><td>{0}</td><td class="mono">{1}</td><td class="mono">{2}</td></tr>' -f (Html-Enc $r.DisplayName), (Html-Enc $r.PrimarySmtpAddress), (Html-Enc $r.Alias))
    }
    $html += '</tbody></table>'
}

# 2.3 Room lists
$html += '<h3>2.3 Room lists</h3>'
$html += ('<p class="small">On-prem source: <span class="mono">{0}</span> &nbsp;&nbsp;|&nbsp;&nbsp; EXO source: live query</p>' -f (Html-Enc (Split-Path -Leaf $OnPremRoomListsCsv)))
$html += '<div class="two-col">'

# On-prem room lists
$html += '<div>'
$html += '<div class="small"><b>On-prem room lists</b></div>'
$html += '<table><thead><tr><th>Display Name</th><th>Primary SMTP</th></tr></thead><tbody>'
foreach ($rl in $onPremRoomLists) {
    $html += ('<tr><td>{0}</td><td class="mono">{1}</td></tr>' -f (Html-Enc $rl.DisplayName), (Html-Enc $rl.PrimarySmtpAddress))
}
$html += '</tbody></table>'
$html += '</div>'

# EXO room lists
$html += '<div>'
$html += '<div class="small"><b>Exchange Online room lists</b></div>'
if (@($exoRoomLists).Count -eq 0) {
    $html += '<p><span class="pill">INFO</span> No room lists were returned from Exchange Online at runtime.</p>'
} else {
    $html += '<table><thead><tr><th>Display Name</th><th>Primary SMTP</th></tr></thead><tbody>'
    foreach ($rl in $exoRoomLists) {
        $html += ('<tr><td>{0}</td><td class="mono">{1}</td></tr>' -f (Html-Enc $rl.DisplayName), (Html-Enc $rl.PrimarySmtpAddress))
    }
    $html += '</tbody></table>'
}
$html += '</div>'

$html += '</div>' # two-col
$html += '</div>' # inventory card

# --- Membership detail section ---
$html += '<div class="card" id="memberships">'
$html += '<h2>3. Room list memberships (rooms per room list)</h2>'
$html += '<p>This section lists the rooms under each room list, showing <b>Display Name</b> and <b>SMTP</b> (On-prem export vs Exchange Online).</p>'

# Build lookups: list smtp -> members objects
$onPremMembersByList = @{}
foreach ($m in $onPremRoomListMembers) {
    $k = ([string]$m.RoomListPrimarySmtpAddress).ToLower()
    if (-not $onPremMembersByList.ContainsKey($k)) { $onPremMembersByList[$k] = @() }
    $onPremMembersByList[$k] += $m
}

$exoMembersByList = @{}
foreach ($m in $exoRoomListMembers) {
    $k = ([string]$m.RoomListPrimarySmtpAddress).ToLower()
    if (-not $exoMembersByList.ContainsKey($k)) { $exoMembersByList[$k] = @() }
    $exoMembersByList[$k] += $m
}

# Union of room lists (prefer on-prem list ordering)
$roomListsForSection = @()
if (@($onPremRoomLists).Count -gt 0) { $roomListsForSection = $onPremRoomLists }
else { $roomListsForSection = $exoRoomLists }

foreach ($rl in $roomListsForSection) {
    $rlSmtp = [string]$rl.PrimarySmtpAddress
    if ([string]::IsNullOrWhiteSpace($rlSmtp)) { continue }
    $k = $rlSmtp.ToLower()

    $opMembers = @()
    if ($onPremMembersByList.ContainsKey($k)) { $opMembers = $onPremMembersByList[$k] }
    $xoMembers = @()
    if ($exoMembersByList.ContainsKey($k)) { $xoMembers = $exoMembersByList[$k] }

    $html += ('<h3>{0} <span class="pill mono">{1}</span></h3>' -f (Html-Enc $rl.DisplayName), (Html-Enc $rlSmtp))
    $html += '<div class="two-col">'

    # On-prem members
    $html += '<div>'
    $html += ('<div class="small"><b>On-prem members</b> &nbsp;<span class="pill">{0}</span></div>' -f @($opMembers).Count)
    if (@($opMembers).Count -eq 0) {
        $html += '<p class="small">No members in on-prem export for this room list.</p>'
    } else {
        $html += '<table><thead><tr><th>Display Name</th><th>SMTP</th><th>Type</th></tr></thead><tbody>'
        foreach ($m in ($opMembers | Sort-Object MemberDisplayName)) {
            $html += ('<tr><td>{0}</td><td class="mono">{1}</td><td>{2}</td></tr>' -f (Html-Enc $m.MemberDisplayName), (Html-Enc $m.MemberPrimarySmtpAddress), (Html-Enc $m.MemberRecipientTypeDetails))
        }
        $html += '</tbody></table>'
    }
    $html += '</div>'

    # EXO members
    $html += '<div>'
    $html += ('<div class="small"><b>EXO members</b> &nbsp;<span class="pill">{0}</span></div>' -f @($xoMembers).Count)
    if (@($xoMembers).Count -eq 0) {
        $html += '<p class="small">No members returned from Exchange Online for this room list at runtime.</p>'
    } else {
        $html += '<table><thead><tr><th>Display Name</th><th>SMTP</th><th>Type</th></tr></thead><tbody>'
        foreach ($m in ($xoMembers | Sort-Object MemberDisplayName)) {
            $html += ('<tr><td>{0}</td><td class="mono">{1}</td><td>{2}</td></tr>' -f (Html-Enc $m.MemberDisplayName), (Html-Enc $m.MemberPrimarySmtpAddress), (Html-Enc $m.MemberRecipientTypeDetails))
        }
        $html += '</tbody></table>'
    }
    $html += '</div>'

    $html += '</div>' # two-col
}

$html += '</div>' # memberships card

# --- Mismatches ---
$html += '<div class="card" id="mismatches">'
$html += '<h2>4. Membership mismatches (On-prem vs EXO)</h2>'
if ($membershipMismatches.Count -eq 0) {
    $html += '<p><span class="pill">PASS</span> No membership mismatches detected.</p>'
} else {
    $html += '<p><span class="pill">ATTENTION</span> One or more room lists have membership differences between on-prem export and Exchange Online.</p>'
    $html += '<table><thead><tr><th>Room List</th><th>Primary SMTP</th><th>Only on-prem</th><th>Only EXO</th></tr></thead><tbody>'
    foreach ($mm in $membershipMismatches) {
        $html += ('<tr><td>{0}</td><td class="mono">{1}</td><td>{2}</td><td>{3}</td></tr>' -f (Html-Enc $mm.RoomListDisplayName), (Html-Enc $mm.RoomListPrimarySmtpAddress), (Html-Enc $mm.OnlyOnPrem), (Html-Enc $mm.OnlyExo))
    }
    $html += '</tbody></table>'
}
$html += '</div>'

# --- Footer ---
$html += '<div class="footer-fixed"><div class="footer-inner">'
$html += ('<div>Generated <span class="mono">{0}</span></div>' -f (Html-Enc $generatedOn))
$html += ('<div><span class="footer-sep">My IT Solutions (PNG) Ltd</span> &middot; By Melky Warinak &middot; <a href="https://myitsolutionspg.com" target="_blank" rel="noopener">myitsolutionspg.com</a></div>')
$html += '</div></div>'

$html += '</div>' # wrap
$html += '</body>'
$html += '</html>'

$htmlText = ($html -join "`r`n")
Set-Content -LiteralPath $outFile -Value $htmlText -Encoding UTF8

Write-Log ("HTML report written to {0}" -f $outFile)

# Disconnect (best-effort)
try { Disconnect-ExchangeOnline -Confirm:$false | Out-Null } catch {}

Write-Log "Rooms & Room Lists migration report completed."
Write-Log "----------------------------------------------------------"