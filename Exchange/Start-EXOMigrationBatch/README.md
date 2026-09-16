# Start-EXOMigrationBatch

Creates Exchange Online migration batches from a list of user principal names.

## Requirements

- Exchange Online PowerShell module version 3 or later
- An active `Connect-ExchangeOnline` session
- A configured migration endpoint and Exchange Administrator permissions

## Key Parameters

- `UserFile`: Required text file containing one user principal name per line.
- `MigrationEndpoint`: Required existing migration endpoint name.
- `TargetDeliveryDomain`: Required Microsoft 365 routing domain.
- `BatchNamePrefix`: Prefix applied to generated batch names.
- `BatchSize`: Users per batch, from `1` through `200`. Defaults to `50`.
- `NotificationEmails`: Addresses that receive migration notifications.
- `AutoStart`: Starts batches after creation.
- `AutoComplete`: Completes batches automatically when ready.
- Common parameter `WhatIf`: Previews batch creation.

## Usage

```powershell
Connect-ExchangeOnline
.\Start-EXOMigrationBatch.ps1 -UserFile C:\Migration\Users.txt `
    -MigrationEndpoint OnpremEndpoint `
    -TargetDeliveryDomain contoso.mail.onmicrosoft.com `
    -BatchSize 25 -AutoStart
```

## Output

The script creates migration batches and writes a processing log and status CSV. Use `-WhatIf` before creating production batches.