# Rooms & Room Lists Migration (Exchange 2016/2019 to Exchange Online)

PowerShell scripts to help migrate **room mailboxes** and **room lists**
from an on-premises **Exchange Server 2016 or Exchange Server 2019** environment to **Exchange Online**.

This repo is designed as a practical toolkit: export on-prem configuration, create
migration batches, (optionally) build cloud room lists, and generate HTML reports for
change records.

![Sample CA report](docs/ca-report.png)

## Supported versions

- **On-prem:** Exchange Server **2016** or **2019**
- **Cloud:** Exchange Online
- **Migration method:** Hybrid **remote move** (Migration Batches)

> Note: This assumes a working Hybrid configuration and a valid Migration Endpoint (SourceEndpoint).
> Keep Exchange servers on supported/current updates in line with Microsoft guidance.

## Contents

- `Rooms-RoomLists/1-Scripts`
  - `1-Export-RoomsAndRoomLists_OnPrem.ps1`  
    Export rooms, calendar settings, room lists and membership from Exchange **2016/2019** to CSV.
  - `2-New-RoomMigrationBatch_Exo.ps1`  
    Create a remote-move migration batch in EXO using `RoomsToMove.csv`.
  - `3-New-ExoRoomLists_FromExport.ps1`  
    (Optional) Create room lists and membership in EXO based on the on-prem CSV exports.
  - `4-New-RoomsAndRoomLists_MigrationReport.ps1`  
    Generate an HTML report comparing on-prem vs EXO rooms and room lists.

## Prerequisites

- **On-prem (Step 1)**
  - Run from **Exchange Management Shell (EMS)** on an **Exchange 2016 or 2019** server  
    (or a management host with the corresponding Exchange management tools).
  - Account must have sufficient permissions to read:
    - room mailboxes and calendar processing settings
    - room list distribution groups and membership

- **Exchange Online (Steps 3–5)**
  - PowerShell module: **ExchangeOnlineManagement**
    ```powershell
    Install-Module ExchangeOnlineManagement -Scope CurrentUser
    ```
  - Account must have permissions to:
    - create and manage migration batches (Hybrid/Move)
    - read and manage distribution groups (if using Step 4)
    - read room mailboxes for reporting

- **Hybrid requirement (for remote move)**
  - A working **Hybrid configuration** and a valid **Migration Endpoint** (the `-SourceEndpoint` value).

## Folder Structure (high level)

```text
Rooms-RoomLists
├─ 1-Scripts
│  ├─ 1-Export-RoomsAndRoomLists_OnPrem.ps1
│  ├─ 2-New-RoomMigrationBatch_Exo.ps1
│  ├─ 3-New-ExoRoomLists_FromExport.ps1
│  └─ 4-New-RoomsAndRoomLists_MigrationReport.ps1
├─ 2-Out          # Reports (generated)
├─ 3-Logs         # Daily script logs (generated)
└─ 4-Export       # On-prem exports + RoomsToMove.csv (generated)

```

## Usage example

### 1. Export rooms and room lists (on-prem)
Run from the Exchange Management Shell on an Exchange 2016 or 2019 server:

```powershell
cd ".\Rooms-RoomLists\1-Scripts" `
.\1-Export-RoomsAndRoomLists_OnPrem.ps1
```
This will:
- Ensure the standard folder structure exists (2-Out, 3-Logs, 4-Export)
- Export the latest CSVs into Rooms-RoomLists/4-Export
- Auto-generate RoomsToMove.csv containing all discovered room mailboxes

### 2. Review / prep `RoomsToMove.csv`

The export script automatically creates `Rooms-RoomLists/4-Export/RoomsToMove.csv`
containing **all** room mailboxes discovered on-prem.

You can either:

- Use this file as-is to migrate every room in a single batch, or  
- Create one or more copies (for example `RoomsToMove_Wave1.csv`, `RoomsToMove_Wave2.csv`)
  and delete rows for rooms that should not be included in that wave.

Each file uses a simple format:

```csv
EmailAddress
room1@contoso.com
room2@contoso.com
room3@contoso.com
```
### 3. Create a migration batch in Exchange Online
Notes:
- AutoStart starts the sync immediately.
- The batch is still manually completed (unless your script explicitly enables auto-complete).

```powershell
cd ".\Rooms-RoomLists\1-Scripts"

.\2-New-RoomMigrationBatch_Exo.ps1 `
  -BatchName "Rooms-To-EXO-Wave1" `
  -SourceEndpoint "Hybrid-MigrationEndpointName" `
  -TargetDeliveryDomain "contoso.mail.onmicrosoft.com" `
  -AutoStart

```
### 4. (Optional) Create room lists in EXO from the export
In many hybrid environments, room lists may already exist in Exchange Online via directory sync.
Only run this step if you are intentionally creating cloud room lists (or need to rebuild them).

```powershell
$include = @(
  "Head Office – Meeting Rooms",
  "Regional Office – Meeting Rooms"
)

.\3-New-ExoRoomLists_FromExport.ps1 `
  -IncludeRoomLists $include `
  -AliasSuffix "-EXO"

```
### 5. Generate an HTML migration report
```powershell
.\4-New-RoomsAndRoomLists_MigrationReport.ps1 `
  -CustomerName "Your Organisation Name" `
  -ProjectName "Exchange Online Migration" `
  -ReportVersion "v1.1"

```
This produces an HTML report under Rooms-RoomLists/2-Out.
