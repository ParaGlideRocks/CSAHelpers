# New-RoomMailboxes

Provision Exchange Online room mailboxes from a semicolon-delimited room inventory CSV, then configure Microsoft Places attributes, calendar processing, and Room Finder room lists.

## Requirements

- PowerShell 5.1 or later
- ExchangeOnlineManagement module V3 or later
- Exchange Online admin permissions to create room mailboxes, update Places metadata, configure calendar processing, and manage room lists
- An active Exchange Online PowerShell session, or credentials that can complete `Connect-ExchangeOnline`

## Files

- `New-RoomMailboxes.ps1`: provisioning script
- `sample-rooms.csv`: sample room inventory using the supported semicolon-delimited format

## Parameters

- `CsvPath`: Room inventory CSV path. Defaults to `.\sample-rooms.csv`.
- `Domain`: SMTP domain used for room mailbox and room list addresses.
- `Prefix`: Optional prefix for room aliases and room list names. A `Prefix` CSV column takes precedence when present.
- `DefaultCountry`: Country code used when the CSV has no country value. Defaults to `IT`.
- `HeaderRow`: Row containing CSV headers. Defaults to `1`.
- `PropagationWaitSeconds`: Wait time after mailbox creation before configuration. Defaults to `120`.
- `MaxAttempts`: Retry attempts while waiting for new mailboxes to become available. Defaults to `6`.
- `RetryDelaySeconds`: Delay between retry attempts. Defaults to `30`.
- `MaxRoomsPerList`: Maximum rooms per room list. Defaults to `50`.
- `SkipRoomLists`: Skips room list creation and membership updates.
- `WhatIfProvisioning`: Runs the workflow in simulation mode without changing the tenant.

## CSV Format

The script accepts both English and Italian column names. The sample uses this schema:

```csv
CountryOrRegion;City;Building;Floor;DisplayName;ResourceCapacity;AudioDeviceName;VideoDeviceName;DisplayDeviceName;Description;Planimetry;RequiresApproval;Approvers
```

Supported fields:

| Field | Supported column names | Notes |
|-------|------------------------|-------|
| Room name | `Nome Sala Riunioni`, `DisplayName`, `Name` | Used as room display name. |
| City | `Sede`, `City` | Used for `Set-Place -City`. |
| Building | `Edificio`, `Building` | Used for `Set-Place -Building` and room list grouping. |
| Floor | `Piano`, `Floor` | Accepts `pt`, `pN`, or an integer. |
| Capacity | `Capienza (numero posti a sedere)`, `Capienza`, `ResourceCapacity`, `Capacity` | Non-numeric values are reduced to the first number and logged as warnings. |
| Audio | `Call conference`, `AudioDeviceName` | Accepts yes/no style values. |
| Video | `Video conference`, `VideoDeviceName` | Accepts yes/no style values. |
| Display | `Proiettore`, `DisplayDeviceName` | Accepts yes/no style values. |
| Description | `Descrizione della Sala *(facoltativo)`, `Descrizione della Sala`, `DeviceDescription` | Optional Places device description. |
| Approval | `Workflow approvativo`, `RequiresApproval` | Enables request-based booking. |
| Approvers | `Approvatore (email)`, `Approvatore`, `Approvers`, `ApproverEmail` | Comma- or semicolon-separated delegates. |
| Country | `CountryOrRegion`, `Country` | Falls back to `DefaultCountry`. |
| Prefix | `Prefix` | Overrides the `Prefix` parameter for that row. |

## Usage

```powershell
Connect-ExchangeOnline

.\New-RoomMailboxes.ps1 -CsvPath .\sample-rooms.csv -Domain contoso.com -WhatIfProvisioning
.\New-RoomMailboxes.ps1 -CsvPath .\rooms.csv -Domain contoso.com
.\New-RoomMailboxes.ps1 -CsvPath .\rooms.csv -Domain contoso.com -SkipRoomLists
```

## Behavior

The script runs in three phases:

1. Creates missing room mailboxes with deterministic aliases.
2. Waits for propagation, then applies mailbox display names, Microsoft Places metadata, device flags, capacity, and calendar processing.
3. Creates room lists grouped by city/building and adds room members, unless `SkipRoomLists` is set.

Provisioning is idempotent: existing rooms are updated rather than recreated, so the script can be rerun after partial failures or propagation delays.

## Output

The script writes:

- `New-RoomMailboxes_<timestamp>.log`
- `New-RoomMailboxes_Report_<timestamp>.csv`

Room Finder and Microsoft Places changes may require 24 to 48 hours to become visible.
