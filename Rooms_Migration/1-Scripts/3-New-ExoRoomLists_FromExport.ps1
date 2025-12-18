<#
.SYNOPSIS
Creates Exchange Online room lists (cloud) based on on-prem exports and adds members.

.FOLDER STRUCTURE (expected)
<ProjectRoot>\
  1-Scripts\
  2-Out\
  3-Logs\
  4-Export\   <-- on-prem CSV exports from Script 1

.DEFAULT CSV SOURCES (latest files in 4-Export)
  RoomLists_OnPrem_*.csv
  RoomListMembers_OnPrem_*.csv

.NOTES
- Requires ExchangeOnlineManagement module.
- By default creates "cloud copy" room lists (AliasSuffix + DisplayNameSuffix) to avoid collisions with DirSync room lists.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$RoomListsCsv,

    [Parameter(Mandatory=$false)]
    [string]$RoomListMembersCsv,

    [Parameter(Mandatory=$false)]
    [string[]]$IncludeRoomLists,   # Match by DisplayName. If not provided, process all.

    [Parameter(Mandatory=$false)]
    [string]$AliasSuffix = "-EXO",

    [Parameter(Mandatory=$false)]
    [string]$DisplayNameSuffix = " (EXO)"
)

# -------------------------
# Paths (relative to 1-Scripts)
# -------------------------
$ScriptRoot   = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot  = Split-Path -Parent $ScriptRoot

$OutPath      = Join-Path $ProjectRoot '2-Out'
$LogsPath     = Join-Path $ProjectRoot '3-Logs'
$ExportPath   = Join-Path $ProjectRoot '4-Export'

foreach ($p in @($OutPath, $LogsPath, $ExportPath)) {
    if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p | Out-Null }
}

$LogFile = Join-Path $LogsPath ("3-New-ExoRoomLists_FromExport_{0}.log" -f (Get-Date -Format "yyyy-MM-dd"))

function Write-Log {
    param([Parameter(Mandatory=$true)][string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    Add-Content -Path $LogFile -Value $line
    Write-Host $line
}

function Get-LatestFile([string]$Folder, [string]$Filter) {
    Get-ChildItem -Path $Folder -Filter $Filter -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

Write-Log "----- EXO Room List Creation script started -----"
Write-Log ("ProjectRoot = {0}" -f $ProjectRoot)
Write-Log ("ExportPath  = {0}" -f $ExportPath)

# Resolve CSV locations
if ([string]::IsNullOrWhiteSpace($RoomListsCsv)) {
    $f = Get-LatestFile -Folder $ExportPath -Filter "RoomLists_OnPrem_*.csv"
    if ($null -eq $f) { throw "RoomListsCsv not provided and no RoomLists_OnPrem_*.csv found in $ExportPath" }
    $RoomListsCsv = $f.FullName
}

if ([string]::IsNullOrWhiteSpace($RoomListMembersCsv)) {
    $f = Get-LatestFile -Folder $ExportPath -Filter "RoomListMembers_OnPrem_*.csv"
    if ($null -eq $f) { throw "RoomListMembersCsv not provided and no RoomListMembers_OnPrem_*.csv found in $ExportPath" }
    $RoomListMembersCsv = $f.FullName
}

Write-Log ("RoomListsCsv       = {0}" -f $RoomListsCsv)
Write-Log ("RoomListMembersCsv = {0}" -f $RoomListMembersCsv)

# Load CSVs
$roomLists = Import-Csv -Path $RoomListsCsv
$members   = Import-Csv -Path $RoomListMembersCsv

if (-not $roomLists -or $roomLists.Count -eq 0) { throw "RoomListsCsv is empty: $RoomListsCsv" }

# Validate expected columns
foreach ($col in @('DisplayName','Alias','PrimarySmtpAddress')) {
    if (-not ($roomLists[0].PSObject.Properties.Name -contains $col)) {
        throw "RoomListsCsv missing required column '$col'. File: $RoomListsCsv"
    }
}

foreach ($col in @('RoomList','MemberPrimarySmtpAddress')) {
    if ($members.Count -gt 0 -and -not ($members[0].PSObject.Properties.Name -contains $col)) {
        throw "RoomListMembersCsv missing required column '$col'. File: $RoomListMembersCsv"
    }
}

Write-Log ("On-prem room lists loaded   = {0}" -f $roomLists.Count)
Write-Log ("On-prem member rows loaded  = {0}" -f ($members.Count))

# Connect to EXO
try {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement module not found. Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop | Out-Null

    Write-Log "Connecting to Exchange Online..."
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
    Write-Log "Connected to Exchange Online."
}
catch {
    Write-Log ("ERROR: Failed to connect to EXO: {0}" -f $_.Exception.Message)
    throw
}

# Filter lists
if ($IncludeRoomLists -and $IncludeRoomLists.Count -gt 0) {
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($n in $IncludeRoomLists) { [void]$set.Add($n) }

    $roomListsToProcess = $roomLists | Where-Object { $set.Contains($_.DisplayName) }
    Write-Log ("Filtering enabled: processing {0} of {1} lists." -f $roomListsToProcess.Count, $roomLists.Count)
} else {
    $roomListsToProcess = $roomLists
    Write-Log "No IncludeRoomLists provided: processing ALL room lists."
}

# Counters
[int]$created = 0
[int]$skippedExisting = 0
[int]$skippedDirSync = 0
[int]$membersAdded = 0
[int]$membersFailed = 0

foreach ($rl in $roomListsToProcess) {

    $srcDisplay = $rl.DisplayName
    $srcAlias   = $rl.Alias
    $srcSmtp    = $rl.PrimarySmtpAddress

    # Domain for primary SMTP (derive from source)
    $domain = ($srcSmtp -split '@')[-1]

    # Target naming (cloud copy)
    $tgtAlias   = "{0}{1}" -f $srcAlias, $AliasSuffix
    $tgtDisplay = "{0}{1}" -f $srcDisplay, $DisplayNameSuffix
    $tgtSmtp    = "{0}@{1}" -f $tgtAlias, $domain

    Write-Log ("--- Room List: {0} ({1}) -> Target: {2} ({3})" -f $srcDisplay, $srcSmtp, $tgtDisplay, $tgtSmtp)

    # Check if target already exists
    $existing = $null
    try {
        $existing = Get-DistributionGroup -Identity $tgtSmtp -ErrorAction SilentlyContinue
    } catch { $existing = $null }

    if ($null -ne $existing) {
        $skippedExisting++

        # If DirSync, do not attempt membership changes
        $isDirSynced = $false
        if ($existing.PSObject.Properties.Name -contains 'IsDirSynced') {
            $isDirSynced = [bool]$existing.IsDirSynced
        }

        if ($isDirSynced) {
            $skippedDirSync++
            Write-Log ("INFO: Target group exists and is DirSync. Skipping membership changes. ({0})" -f $tgtSmtp)
            continue
        }

        Write-Log ("INFO: Target group already exists (cloud). Will attempt to ensure membership. ({0})" -f $tgtSmtp)
        $targetGroup = $existing
    }
    else {
        # Create new RoomList group (cloud)
        try {
            $targetGroup = New-DistributionGroup -Name $tgtDisplay -Alias $tgtAlias -RoomList -ErrorAction Stop

            # Ensure primary SMTP is as expected
            try {
                Set-DistributionGroup -Identity $targetGroup.Identity -PrimarySmtpAddress $tgtSmtp -ErrorAction Stop | Out-Null
            } catch {}

            $created++
            Write-Log ("Created room list: {0} ({1})" -f $tgtDisplay, $tgtSmtp)
        }
        catch {
            Write-Log ("ERROR: Failed to create room list '{0}': {1}" -f $tgtDisplay, $_.Exception.Message)
            continue
        }
    }

    # Add members (based on source list membership)
    $listMembers = $members | Where-Object {
        $_.RoomList -and ($_.RoomList.Trim().ToLower() -eq $srcSmtp.Trim().ToLower())
    }

    if (-not $listMembers -or $listMembers.Count -eq 0) {
        Write-Log "No members found in export for this room list."
        continue
    }

    foreach ($m in $listMembers) {
        $memberSmtp = ($m.MemberPrimarySmtpAddress | ForEach-Object { $_.Trim() })

        if ([string]::IsNullOrWhiteSpace($memberSmtp)) { continue }

        # Confirm member exists in EXO directory (recipient)
        $r = $null
        try { $r = Get-Recipient -Identity $memberSmtp -ErrorAction SilentlyContinue } catch { $r = $null }

        if ($null -eq $r) {
            $membersFailed++
            Write-Log ("WARN: Member not found in EXO directory, skipping: {0}" -f $memberSmtp)
            continue
        }

        # Add member (ignore if already member)
        try {
            Add-DistributionGroupMember -Identity $targetGroup.Identity -Member $memberSmtp -ErrorAction Stop | Out-Null
            $membersAdded++
            Write-Log ("Added member: {0}" -f $memberSmtp)
        }
        catch {
            # If already a member, EXO usually throws; treat as info
            $msg = $_.Exception.Message
            if ($msg -match "is already a member" -or $msg -match "already a member") {
                Write-Log ("INFO: Already a member: {0}" -f $memberSmtp)
            } else {
                $membersFailed++
                Write-Log ("WARN: Failed to add member {0}: {1}" -f $memberSmtp, $msg)
            }
        }
    }
}

Write-Log "----- Summary -----"
Write-Log ("Room lists processed       = {0}" -f $roomListsToProcess.Count)
Write-Log ("Room lists created         = {0}" -f $created)
Write-Log ("Room lists existed (skip)  = {0}" -f $skippedExisting)
Write-Log ("DirSync lists skipped      = {0}" -f $skippedDirSync)
Write-Log ("Members added              = {0}" -f $membersAdded)
Write-Log ("Members failed/skipped     = {0}" -f $membersFailed)

Write-Log "----- Script completed -----"
