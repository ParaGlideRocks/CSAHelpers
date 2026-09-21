# Export-DmarcAttachments

Saves DMARC aggregate report attachments from an Outlook folder to disk, ready for `Analyze-DmarcReports.ps1`.

## Requirements

- PowerShell 5.1 or later, running **without** elevation
- Outlook desktop client (classic), running and signed in
- Full Access permission on the mailbox holding the reports

## Parameters

- `FolderPath`: Folder to export, as a backslash-separated path from the mailbox root. Outlook prompts for the folder when omitted.
- `OutputPath`: Destination directory. Defaults to `.\DmarcReports`.
- `Since`: Only exports messages received on or after this date.
- `Until`: Only exports messages received before this date.
- `Extensions`: Attachment extensions to export. Defaults to `gz`, `zip`, and `xml`.
- `MoveToFolder`: Existing folder the exported messages are moved to.
- `IncludeSubfolders`: Also processes the subfolders of the selected folder.
- `Overwrite`: Re-exports attachments that are already on disk.

## Usage

```powershell
.\Export-DmarcAttachments.ps1
.\Export-DmarcAttachments.ps1 -FolderPath "dmarc_rua@contoso.com\Inbox" -OutputPath C:\DMARC\raw
.\Export-DmarcAttachments.ps1 -FolderPath "dmarc_rua@contoso.com\Inbox" -Since 2026-08-01 -WhatIf
```

## Output

Attachments are written to the destination directory, each prefixed with the timestamp of the message that carried it, because reporting providers reuse the same attachment names and would otherwise overwrite one another. Files already present are skipped, so the export can be re-run incrementally to collect only what has arrived since the previous run. A summary of messages examined, attachments exported, skipped, and failed is printed at the end.

## Notes

Outlook automation requires this script and Outlook to run at the same Windows integrity level. A session started with **Run as administrator** cannot attach to a normally started Outlook and fails with `CO_E_SERVER_EXEC_FAILURE` (0x80080005); the script checks for this before doing any work. The new Outlook for Windows does not expose COM automation at all, and a mailbox in another tenant cannot be reached this way.

Messages are enumerated in reverse order and restricted by date through MAPI rather than filtered afterwards, so that folders holding tens of thousands of reports remain workable and `MoveToFolder` does not disturb the iteration.

## Related Script

Use `Analyze-DmarcReports.ps1` to parse the exported reports and produce the summary workbook.
