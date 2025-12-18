<#
.SYNOPSIS
  Creates an Exchange Online remote-move Migration Batch for room mailboxes using RoomsToMove.csv.

.NOTES
  Folder layout expected (relative to script folder):
    ..\2-Out
    ..\3-Logs
    ..\4-Export

  IMPORTANT:
  - New-MigrationBatch (remote move/onboarding) does NOT support -BadItemLimit on many tenants/module builds.
    If a mailbox fails with TooManyBadItems, adjust per-move-request later using Set-MoveRequest.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$BatchName,

  [Parameter(Mandatory = $true)]
  [string]$SourceEndpoint,

  [Parameter(Mandatory = $true)]
  [string]$TargetDeliveryDomain,

  # Optional: point to a wave file like RoomsToMove_Wave1.csv
  [Parameter(Mandatory = $false)]
  [string]$RoomsCsv,

  [switch]$AutoStart
)

# ---------------------------
# Paths + logging helpers
# ---------------------------
$ScriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir

$OutPath    = Join-Path $ProjectRoot "2-Out"
$LogsPath   = Join-Path $ProjectRoot "3-Logs"
$ExportPath = Join-Path $ProjectRoot "4-Export"

foreach ($p in @($OutPath,$LogsPath,$ExportPath)) {
  if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p | Out-Null }
}

if (-not $RoomsCsv) {
  $RoomsCsv = Join-Path $ExportPath "RoomsToMove.csv"
}

$LogFile = Join-Path $LogsPath ("2-New-RoomMigrationBatch_Exo_{0}.log" -f (Get-Date -Format "yyyy-MM-dd_HHmmss"))

function Write-Log {
  param([string]$Message)
  $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  $line = "[{0}] {1}" -f $ts, $Message
  Write-Host $line
  Add-Content -Path $LogFile -Value $line
}

Write-Log "----- Room Migration Batch script started -----"
Write-Log "ProjectRoot      = $ProjectRoot"
Write-Log "ExportPath       = $ExportPath"
Write-Log "OutPath          = $OutPath"
Write-Log "LogsPath         = $LogsPath"
Write-Log "RoomsCsv         = $RoomsCsv"

# ---------------------------
# Validate CSV + build CSVData
# ---------------------------
if (-not (Test-Path $RoomsCsv)) {
  throw "Rooms CSV not found: $RoomsCsv"
}

$rows = Import-Csv -Path $RoomsCsv
if (-not $rows -or $rows.Count -eq 0) {
  throw "Rooms CSV is empty: $RoomsCsv"
}

# Expect EmailAddress column (preferred). If not present, try common alternates.
$first = $rows | Select-Object -First 1
$cols = $first.PSObject.Properties.Name

$emailCol = $null
foreach ($c in @("EmailAddress","PrimarySmtpAddress","UserPrincipalName","UPN","SmtpAddress")) {
  if ($cols -contains $c) { $emailCol = $c; break }
}
if (-not $emailCol) {
  throw "Rooms CSV must contain an EmailAddress column (or one of: PrimarySmtpAddress, UserPrincipalName, UPN, SmtpAddress). Found: $($cols -join ', ')"
}

$emails = $rows |
  ForEach-Object { ($_.$emailCol | ForEach-Object { "$_".Trim() }) } |
  Where-Object { $_ -and $_ -match "@" } |
  Sort-Object -Unique

Write-Log ("Rooms in CSV (unique) = {0}" -f $emails.Count)

if ($emails.Count -eq 0) {
  throw "No valid email addresses found in $RoomsCsv (column: $emailCol)."
}

# Build minimal CSV payload required by New-MigrationBatch
$csvLines = @("EmailAddress")
$csvLines += $emails
$csvString = ($csvLines -join "`r`n") + "`r`n"
$csvBytes  = [System.Text.Encoding]::UTF8.GetBytes($csvString)

# Save the exact payload used (for traceability)
$payloadPath = Join-Path $OutPath ("RoomsToMove_Payload_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
Set-Content -Path $payloadPath -Value $csvString -Encoding UTF8
Write-Log "Saved payload CSV used for batch: $payloadPath"

# ---------------------------
# Connect to EXO (assumes ExchangeOnlineManagement is installed)
# ---------------------------
Write-Log "Connecting to Exchange Online..."
try {
  if (-not (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
    throw "Connect-ExchangeOnline not found. Install-Module ExchangeOnlineManagement and import it."
  }

  # If already connected, this is harmless.
  Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop | Out-Null
  Write-Log "Connected to Exchange Online."
}
catch {
  Write-Log "ERROR: Failed to connect to Exchange Online: $($_.Exception.Message)"
  throw
}

# ---------------------------
# Create migration batch
# ---------------------------
try {
  Write-Log "Creating migration batch: $BatchName"

  $params = @{
    Name               = $BatchName
    SourceEndpoint     = $SourceEndpoint
    TargetDeliveryDomain = $TargetDeliveryDomain
    CSVData            = $csvBytes
    ErrorAction        = 'Stop'
  }

  if ($AutoStart.IsPresent) {
    $params.AutoStart = $true
  }

  $newBatch = New-MigrationBatch @params

  Write-Log "Migration batch created successfully."
  Write-Log ("Batch Name  : {0}" -f $newBatch.Identity)
  Write-Log ("Status      : {0}" -f $newBatch.Status)
  Write-Log ("State       : {0}" -f $newBatch.State)

  # Dump key details to an output file
  $outFile = Join-Path $OutPath ("MigrationBatch_{0}_{1}.txt" -f ($BatchName -replace '[^\w\- ]',''), (Get-Date -Format "yyyyMMdd_HHmmss"))
  $newBatch | Format-List * | Out-String | Set-Content -Path $outFile -Encoding UTF8
  Write-Log "Saved batch details: $outFile"
}
catch {
  Write-Log "ERROR: Failed to create migration batch: $($_.Exception.Message)"
  throw
}
finally {
  Write-Log "----- Script completed -----"
}
