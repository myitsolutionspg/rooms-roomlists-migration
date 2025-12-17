<#
.SYNOPSIS
Create a remote-move migration batch for room mailboxes.

DEFAULT RoomsCsv:
  ..\2-Out\RoomsToMove.csv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RoomsCsv,                 # Defaults to ..\2-Out\RoomsToMove.csv

    [Parameter(Mandatory = $true)]
    [string]$BatchName,

    [Parameter(Mandatory = $true)]
    [string]$SourceEndpoint,

    [Parameter(Mandatory = $true)]
    [string]$TargetDeliveryDomain,

    [switch]$AutoStart,
    [switch]$AutoComplete
)

#region Paths + logging setup
$scriptRoot         = Split-Path -Parent $MyInvocation.MyCommand.Path     # ...\Rooms_Migration\1-Scripts
$roomsMigrationRoot = Split-Path -Parent $scriptRoot                      # ...\Rooms_Migration
$projectRoot        = Split-Path -Parent $roomsMigrationRoot              # ...\ROOMS AND ROOM LIST MIGRATION
$outRoot            = Join-Path $roomsMigrationRoot '2-Out'
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

#region Resolve default RoomsCsv
if (-not $RoomsCsv) {
    $RoomsCsv = Join-Path $outRoot 'RoomsToMove.csv'
}
#endregion

Write-Log "----- Room migration batch script started -----"
Write-Log "RoomsCsv = $RoomsCsv; BatchName = $BatchName; Endpoint = $SourceEndpoint; TargetDomain = $TargetDeliveryDomain"

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

#region Validate CSV
if (-not (Test-Path -Path $RoomsCsv)) {
    Write-Log "RoomsCsv not found: $RoomsCsv" -Level ERROR
    throw "RoomsCsv not found."
}

Write-Host "Using room CSV: $RoomsCsv" -ForegroundColor Cyan
$roomsList = Import-Csv -Path $RoomsCsv
Write-Log "Rooms in CSV = $($roomsList.Count)"

$csvBytes  = [System.IO.File]::ReadAllBytes($RoomsCsv)
#endregion

#region Create migration batch
$batchParams = @{
    Name                 = $BatchName
    SourceEndpoint       = $SourceEndpoint
    CSVData              = $csvBytes
    TargetDeliveryDomain = $TargetDeliveryDomain
}

if ($AutoStart)    { $batchParams.Add("AutoStart",    $true) }
if ($AutoComplete) { $batchParams.Add("AutoComplete", $true) }

Write-Host "Creating migration batch '$BatchName'..." -ForegroundColor Yellow
$newBatch = New-MigrationBatch @batchParams

Write-Host "Migration batch created:" -ForegroundColor Green
$newBatch | Format-List Name,Status,TotalCount,SourceEndpoint,TargetDeliveryDomain

Write-Log ("Migration batch '{0}' created. Status = {1}; TotalCount = {2}" -f `
           $newBatch.Name, $newBatch.Status, $newBatch.TotalCount)
#endregion

Write-Host "`nNext steps:" -ForegroundColor Cyan
Write-Host "  Start:   Start-MigrationBatch -Identity '$BatchName'"
Write-Host "  Monitor: Get-MigrationBatch -Identity '$BatchName' | fl Status,TotalCount,CompleteCount"
Write-Host "  Users:   Get-MigrationUser -BatchId '$BatchName' | ft Identity,Status,ItemsSynced,TotalItemsInSourceMailbox"

Write-Log "Room migration batch script completed."
Write-Log "-------------------------------------------"
