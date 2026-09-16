# Remove-InvalidSMTP

Removes secondary SMTP proxy addresses that belong to specified invalid domains from Exchange mailboxes.

## Requirements

- Exchange Online or on-premises Exchange cmdlets
- An authenticated session with permission to modify mailboxes

## Parameters

- `UserFile`: Required text file containing one user principal name per line.
- `InvalidDomains`: Required list of domains whose secondary addresses should be removed.
- `LogFile`: Destination log path. Defaults to a timestamped file.
- Common parameter `WhatIf`: Previews changes without applying them.

## Usage

```powershell
.\Remove-InvalidSMTP.ps1 -UserFile C:\Data\Users.txt -InvalidDomains olddomain.com,legacy.org -WhatIf
```

## Warning

This script modifies production mailboxes. Primary SMTP addresses are preserved, but email address policy processing can be disabled for affected mailboxes. Review `-WhatIf` output before running without it.