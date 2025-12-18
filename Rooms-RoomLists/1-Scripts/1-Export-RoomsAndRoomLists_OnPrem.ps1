<#
.SYNOPSIS
Export on-premises rooms, calendar settings, room lists and membership.

Also ensures the standard folder structure exists and auto-generates a
RoomsToMove.csv file (all rooms) under 4-Export.

Folder layout (relative to this script):
  1-Scripts   (this script lives here)
  2-Out       (used by other scripts)
  3-Logs      (log files)
  4-Export    (raw CSV exports from on-prem, incl. RoomsToMove.csv)

Run this script from the Exchange Management Shell on an Exchange 2016 server.
#>

[CmdletBinding()]
param(
    # Optional override for where CSV exports are written.
    [string]$ExportPath
)

$ErrorActionPreference = 'Stop'

# Resolve base paths from script location
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path   # ...\Rooms-RoomLists\1-Scripts
$rootPath   = Split-Path $scriptRoot -Parent                    # ...\Rooms-RoomLists

$paths = [ordered]@{
    Scripts = $scriptRoot
    Out     = Join-Path $rootPath '2-Out'
    Logs    = Join-Path $rootPath '3-Logs'
    Export  = if ($ExportPath) { $ExportPath } else { Join-Path $rootPath '4-Export' }
}

# Ensure folders exist (except 1-Scripts which already exists)
foreach ($key in $paths.Keys) {
    if ($key -eq 'Scripts') { continue }
    if (-not (Test-Path $paths[$key])) {
        New-Item -ItemType Directory -Path $paths[$key] -Force | Out-Null
    }
}

$exportPath = $paths.Export
$logPath    = $paths.Logs

$timestamp = Get-Date -Format 'yyyyMMdd_HHmm'
$today     = Get-Date -Format 'yyyy-MM-dd'
$logFile   = Join-Path $logPath ("1-Export-RoomsAndRoomLists_OnPrem_{0}.log" -f $today)

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line  = "[{0}] {1}" -f $stamp, $Message
    Write-Host $Message
    Add-Content -Path $logFile -Value $line
}

Write-Log "----- Rooms & Room Lists export started -----"
Write-Log "ExportPath = $exportPath"

try {
    # ----------------------- Rooms -----------------------
    Write-Log "Exporting room mailboxes..."
    $rooms = Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize Unlimited |
             Select-Object `
                DisplayName,
                Alias,
                PrimarySmtpAddress,
                WindowsEmailAddress,
                OrganizationalUnit,
                Office,
                ResourceCapacity,
                CustomAttribute1,
                CustomAttribute2,
                CustomAttribute3

    $roomsCsv = Join-Path $exportPath ("Rooms_OnPrem_{0}.csv" -f $timestamp)
    $rooms | Export-Csv -Path $roomsCsv -NoTypeInformation -Encoding UTF8
    Write-Log ("Rooms exported to {0} (Count = {1})" -f $roomsCsv, ($rooms.Count))

    # ---------------- Calendar Processing ---------------
    Write-Log "Exporting CalendarProcessing settings..."
        $calData = foreach ($room in $rooms) {
        $id = $room.PrimarySmtpAddress.ToString()   # force to string
        $cp = Get-CalendarProcessing -Identity $id
    
        [PSCustomObject]@{
            DisplayName                    = $room.DisplayName
            PrimarySmtpAddress             = $room.PrimarySmtpAddress
            AutomateProcessing             = $cp.AutomateProcessing
            BookingWindowInDays            = $cp.BookingWindowInDays
            MaximumDurationInMinutes       = $cp.MaximumDurationInMinutes
            AllowRecurringMeetings         = $cp.AllowRecurringMeetings
            EnforceCapacity                = $cp.EnforceCapacity
            AllowBookingWithoutOrganizer   = $cp.AllowBookingWithoutOrganizer
            AllowOutOfOffice               = $cp.AllowOutOfOffice
            AllBookInPolicy                = $cp.AllBookInPolicy
            AllRequestInPolicy             = $cp.AllRequestInPolicy
            AllRequestOutOfPolicy          = $cp.AllRequestOutOfPolicy
        }
    }

    $calCsv = Join-Path $exportPath ("Rooms_Calendar_OnPrem_{0}.csv" -f $timestamp)
    $calData | Export-Csv -Path $calCsv -NoTypeInformation -Encoding UTF8
    Write-Log ("Calendar settings exported to {0} (Rooms with settings = {1})" -f $calCsv, ($calData.Count))

    # -------------------- Room Lists --------------------
    Write-Log "Exporting room lists..."
    $roomLists = Get-DistributionGroup -RecipientTypeDetails RoomList -ResultSize Unlimited |
             Select-Object `
                Identity,                                      # keep full identity
                DisplayName,
                Alias,
                @{ Name = 'PrimarySmtpAddress'; Expression = { $_.PrimarySmtpAddress.ToString() } },
                ManagedBy,
                Notes

    $roomListsCsv = Join-Path $exportPath ("RoomLists_OnPrem_{0}.csv" -f $timestamp)
    $roomLists | Export-Csv -Path $roomListsCsv -NoTypeInformation -Encoding UTF8
    Write-Log ("Room lists exported to {0} (Count = {1})" -f $roomListsCsv, ($roomLists.Count))

    # ---------------- Room List Membership --------------
    Write-Log "Exporting room list membership..."

    $listMembers = foreach ($list in $roomLists) {
    
        $id = $list.Identity.ToString()
        Write-Log ("  Getting members for room list: {0} [{1}]" -f $list.DisplayName, $id)
    
        try {
            $members = Get-DistributionGroupMember -Identity $id -ResultSize Unlimited -ErrorAction Stop
        }
        catch {
            Write-Log ("  WARN: Failed to get members for room list '{0}' ({1}): {2}" -f $list.DisplayName, $id, $_.Exception.Message)
            continue
        }
    
        foreach ($m in $members) {
            [PSCustomObject]@{
                RoomListDisplayName        = $list.DisplayName
                RoomListPrimarySmtpAddress = $list.PrimarySmtpAddress
                MemberDisplayName          = $m.DisplayName
                MemberPrimarySmtpAddress   = $m.PrimarySmtpAddress
                MemberRecipientType        = $m.RecipientType
                MemberRecipientTypeDetails = $m.RecipientTypeDetails
            }
        }
    }
    
    $listMembersCsv = Join-Path $exportPath ("RoomListMembers_OnPrem_{0}.csv" -f $timestamp)
    $listMembers | Export-Csv -Path $listMembersCsv -NoTypeInformation -Encoding UTF8
    Write-Log ("Room list members exported to {0} (Rows = {1})" -f $listMembersCsv, ($listMembers.Count))

    # ----------------- RoomsToMove.csv ------------------
    Write-Log "Generating RoomsToMove.csv with all rooms..."
    $roomsToMoveCsv = Join-Path $exportPath 'RoomsToMove.csv'

    $rooms |
      Select-Object @{Name = 'EmailAddress'; Expression = { $_.PrimarySmtpAddress }} |
      Export-Csv -Path $roomsToMoveCsv -NoTypeInformation -Encoding UTF8

    Write-Log ("RoomsToMove.csv generated at {0} (Count = {1})" -f $roomsToMoveCsv, ($rooms.Count))

    Write-Log "Completed Rooms & Room Lists export."
    Write-Log "----------------------------------------------------------"
}
catch {
    Write-Log ("ERROR: {0}" -f $_.Exception.Message)
    throw
}
