<#
.SYNOPSIS
Create EXO room lists from on-prem exports.

DEFAULTS:
  RoomListsCsv        = latest RoomLists_OnPrem_*.csv in ..\..\Export CSV
  RoomListMembersCsv  = latest RoomListMembers_OnPrem_*.csv in ..\..\Export CSV
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$RoomListsCsv,

    [Parameter(Mandatory = $false)]
    [string]$RoomListMembersCsv,

    [Parameter(Mandatory = $false)]
    [string[]]$IncludeRoomLists,  # DisplayName OR PrimarySmtpAddress

    [Parameter(Mandatory = $false)]
    [string]$AliasSuffix = "-EXO"
)

#region Paths + logging setup
$scriptRoot         = Split-Path -Parent $MyInvocation.MyCommand.Path     # ...\Rooms_Migration\1-Scripts
$roomsMigrationRoot = Split-Path -Parent $scriptRoot                      # ...\Rooms_Migration
$projectRoot        = Split-Path -Parent $roomsMigrationRoot              # ...\ROOMS AND ROOM LIST MIGRATION
$exportCsvRoot      = Join-Path $projectRoot 'Export CSV'
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

#region Resolve default CSV paths
if (-not $RoomListsCsv) {
    $latest = Get-ChildItem -Path $exportCsvRoot -Filter 'RoomLists_OnPrem_*.csv' -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latest) { $RoomListsCsv = $latest.FullName }
}
if (-not $RoomListMembersCsv) {
    $latest = Get-ChildItem -Path $exportCsvRoot -Filter 'RoomListMembers_OnPrem_*.csv' -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latest) { $RoomListMembersCsv = $latest.FullName }
}
#endregion

Write-Log "----- EXO Room List creation started -----"
Write-Log "RoomListsCsv = $RoomListsCsv"
Write-Log "RoomListMembersCsv = $RoomListMembersCsv"
Write-Log "AliasSuffix = $AliasSuffix"

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

#region Load CSVs
if (-not $RoomListsCsv -or -not (Test-Path -Path $RoomListsCsv)) {
    Write-Log "RoomListsCsv not found. Expected RoomLists_OnPrem_*.csv in '$exportCsvRoot' or pass -RoomListsCsv." -Level ERROR
    throw "RoomListsCsv not found."
}
if (-not $RoomListMembersCsv -or -not (Test-Path -Path $RoomListMembersCsv)) {
    Write-Log "RoomListMembersCsv not found. Expected RoomListMembers_OnPrem_*.csv in '$exportCsvRoot' or pass -RoomListMembersCsv." -Level ERROR
    throw "RoomListMembersCsv not found."
}

$roomLists  = Import-Csv -Path $RoomListsCsv
$membersCsv = Import-Csv -Path $RoomListMembersCsv

if ($IncludeRoomLists) {
    $roomLists = $roomLists | Where-Object {
        $IncludeRoomLists -contains $_.DisplayName -or
        $IncludeRoomLists -contains $_.PrimarySmtpAddress
    }
}

if (-not $roomLists) {
    Write-Log "No room lists selected after filtering. Nothing to do." -Level WARN
    return
}

Write-Host "Room lists selected for EXO creation:" -ForegroundColor Cyan
$roomLists | Select-Object DisplayName,PrimarySmtpAddress,Alias | Format-Table

Write-Log "Room lists selected for creation = $($roomLists.Count)"
#endregion

#region Create EXO room lists
$listsCreated      = 0
$listsSkipped      = 0
$totalMembersAdded = 0

foreach ($rl in $roomLists) {

    $origAlias   = $rl.Alias
    $origSmtp    = $rl.PrimarySmtpAddress
    $displayName = $rl.DisplayName

    if (-not $origAlias) {
        $origAlias = ($displayName -replace '\s','') -replace '[^a-zA-Z0-9]',''
    }

    $newAlias       = "$origAlias$AliasSuffix"
    $domain         = ($origSmtp -split '@')[1]
    $newPrimarySmtp = "$newAlias@$domain"

    $aliasExists = Get-Recipient -Filter "Alias -eq '$newAlias'" -ErrorAction SilentlyContinue
    $smtpExists  = Get-Recipient -Filter "EmailAddresses -eq 'SMTP:$newPrimarySmtp'" -ErrorAction SilentlyContinue

    if ($aliasExists -or $smtpExists) {
        $listsSkipped++
        Write-Warning "Skipping '$displayName' - alias or SMTP already exists (Alias=$newAlias, SMTP=$newPrimarySmtp)"
        Write-Log "Skipping '$displayName' - alias or SMTP already exists (Alias=$newAlias, SMTP=$newPrimarySmtp)" -Level WARN
        continue
    }

    $what = "Create EXO room list '$displayName' (Alias=$newAlias, PrimarySmtp=$newPrimarySmtp)"

    if ($PSCmdlet.ShouldProcess($displayName, $what)) {

        Write-Host $what -ForegroundColor Yellow

        $newGroup = New-DistributionGroup `
            -Name $displayName `
            -DisplayName $displayName `
            -Alias $newAlias `
            -PrimarySmtpAddress $newPrimarySmtp `
            -RoomList

        $listsCreated++
        $membersAddedForList = 0

        Write-Log "Created EXO room list '$displayName' (Alias=$newAlias, PrimarySmtp=$newPrimarySmtp)"

        if ($rl.HiddenFromAddressListsEnabled -ne $null) {
            Set-DistributionGroup $newGroup.Identity -HiddenFromAddressListsEnabled:$rl.HiddenFromAddressListsEnabled
        }
        if ($rl.RequireSenderAuthenticationEnabled -ne $null) {
            Set-DistributionGroup $newGroup.Identity -RequireSenderAuthenticationEnabled:$rl.RequireSenderAuthenticationEnabled
        }

        $rlMembers = $membersCsv | Where-Object { $_.RoomList -eq $origSmtp }

        foreach ($m in $rlMembers) {
            $memberSmtp = $m.MemberPrimarySmtp
            if (-not $memberSmtp) {
                Write-Warning "  Skipping member with no SMTP for room list $displayName"
                Write-Log      "Skipping member with no SMTP for room list $displayName" -Level WARN
                continue
            }

            $recipient = Get-Recipient -Identity $memberSmtp -ErrorAction SilentlyContinue
            if ($recipient) {
                Add-DistributionGroupMember -Identity $newGroup.Identity -Member $recipient.Identity
                $membersAddedForList++
                $totalMembersAdded++
            }
            else {
                Write-Warning "  ! Could not resolve member $memberSmtp for room list $displayName"
                Write-Log      "Could not resolve member $memberSmtp for room list $displayName" -Level WARN
            }
        }

        Write-Log "Room list '$displayName' members added = $membersAddedForList"
    }
}
#endregion

Write-Host "`nCompleted EXO room list creation." -ForegroundColor Cyan
Write-Host "Verify with: Get-DistributionGroup -RecipientTypeDetails RoomList | ft DisplayName,PrimarySmtpAddress,Alias"

Write-Log "EXO Room List creation completed. Created = $listsCreated; Skipped = $listsSkipped; MembersAddedTotal = $totalMembersAdded."
Write-Log "-------------------------------------------------------------"
