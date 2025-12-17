<#
.SYNOPSIS
Export rooms and room lists from on-prem Exchange 2016.

DEFAULT OUTPUT:
  ..\..\Export CSV\
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath
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

#region Resolve default OutputPath
if (-not $OutputPath) {
    $OutputPath = $exportCsvRoot
}
#endregion

#region Prep
if (-not (Test-Path -Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$timestamp           = Get-Date -Format "yyyyMMdd_HHmm"
$roomsFile           = Join-Path $OutputPath "Rooms_OnPrem_$timestamp.csv"
$roomsCalFile        = Join-Path $OutputPath "Rooms_Calendar_OnPrem_$timestamp.csv"
$roomListsFile       = Join-Path $OutputPath "RoomLists_OnPrem_$timestamp.csv"
$roomListMembersFile = Join-Path $OutputPath "RoomListMembers_OnPrem_$timestamp.csv"

Write-Host "Export path: $OutputPath" -ForegroundColor Cyan
Write-Log  "----- Rooms & Room Lists export started -----"
Write-Log  "OutputPath = $OutputPath"
#endregion

#region Export room mailboxes
Write-Host "Exporting room mailboxes..." -ForegroundColor Yellow

$rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited

$rooms |
    Select-Object `
        DisplayName,
        Alias,
        PrimarySmtpAddress,
        LegacyExchangeDN,
        Database,
        OrganizationalUnit,
        RecipientTypeDetails,
        ResourceCapacity,
        Office,
        CustomAttribute1,
        CustomAttribute2,
        CustomAttribute3 |
    Export-Csv -Path $roomsFile -NoTypeInformation -Encoding UTF8

Write-Host "  -> Rooms exported to $roomsFile" -ForegroundColor Green
Write-Log  "Rooms exported to $roomsFile (Count = $($rooms.Count))"
#endregion

#region Export CalendarProcessing
Write-Host "Exporting CalendarProcessing settings..." -ForegroundColor Yellow

$calSettings = foreach ($rm in $rooms) {
    try {
        Get-CalendarProcessing -Identity $rm.Identity -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to get CalendarProcessing for $($rm.Identity): $($_.Exception.Message)"
        Write-Log "Failed to get CalendarProcessing for $($rm.Identity): $($_.Exception.Message)" -Level WARN
    }
}

$calSettings |
    Select-Object `
        Identity,
        AutomateProcessing,
        BookingWindowInDays,
        AllowConflicts,
        AllowRecurringMeetings,
        EnforceSchedulingHorizon,
        MaximumDurationInMinutes,
        DeleteSubject,
        DeleteComments,
        AddOrganizerToSubject,
        RemovePrivateProperty,
        ForwardRequestsToDelegates,
        AllBookInPolicy,
        AllRequestInPolicy,
        AllRequestOutOfPolicy |
    Export-Csv -Path $roomsCalFile -NoTypeInformation -Encoding UTF8

Write-Host "  -> Calendar settings exported to $roomsCalFile" -ForegroundColor Green
Write-Log  "CalendarProcessing exported to $roomsCalFile (Rooms with settings = $($calSettings.Count))"
#endregion

#region Export room lists
Write-Host "Exporting room lists..." -ForegroundColor Yellow

$roomLists = Get-DistributionGroup -RecipientTypeDetails RoomList -ResultSize Unlimited

$roomLists |
    Select-Object `
        DisplayName,
        Alias,
        PrimarySmtpAddress,
        LegacyExchangeDN,
        ManagedBy,
        HiddenFromAddressListsEnabled,
        ModerationEnabled,
        BypassModerationFromSendersOrMembers,
        AcceptMessagesOnlyFromSendersOrMembers,
        RejectMessagesFromSendersOrMembers,
        RequireSenderAuthenticationEnabled,
        CustomAttribute1,
        CustomAttribute2,
        CustomAttribute3 |
    Export-Csv -Path $roomListsFile -NoTypeInformation -Encoding UTF8

Write-Host "  -> Room lists exported to $roomListsFile" -ForegroundColor Green
Write-Log  "Room lists exported to $roomListsFile (Count = $($roomLists.Count))"
#endregion

#region Export room list membership
Write-Host "Exporting room list membership..." -ForegroundColor Yellow

$roomListMembers = foreach ($rl in $roomLists) {
    $members = Get-DistributionGroupMember -Identity $rl.Identity -ResultSize Unlimited
    foreach ($m in $members) {
        [PSCustomObject]@{
            RoomList               = $rl.PrimarySmtpAddress
            RoomListDisplayName    = $rl.DisplayName
            MemberDisplayName      = $m.DisplayName
            MemberPrimarySmtp      = $m.PrimarySmtpAddress
            MemberAlias            = $m.Alias
            MemberRecipientType    = $m.RecipientType
            MemberRecipientDetails = $m.RecipientTypeDetails
        }
    }
}

$roomListMembers |
    Export-Csv -Path $roomListMembersFile -NoTypeInformation -Encoding UTF8

Write-Host "  -> Room list membership exported to $roomListMembersFile" -ForegroundColor Green
Write-Log  "Room list members exported to $roomListMembersFile (Rows = $($roomListMembers.Count))"
#endregion

Write-Host "`nCompleted Rooms & Room Lists export." -ForegroundColor Cyan
Write-Host "Files generated in: $OutputPath"

Write-Log "Export completed successfully."
Write-Log "-------------------------------------------"
