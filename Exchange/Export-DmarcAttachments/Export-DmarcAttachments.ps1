<#
.SYNOPSIS
    Saves DMARC aggregate report attachments from an Outlook folder to disk.

.DESCRIPTION
    Connects to the running Outlook desktop client and exports the attachments
    of the messages held in a mail folder, ready to be processed by
    Analyze-DmarcReports.ps1.

    The script is designed for RUA mailboxes, which typically receive one
    compressed XML per reporting provider per day:

      - restricts the messages server-side by date, so large folders stay usable
      - prefixes every file with the message timestamp, because providers reuse
        the same attachment names and would otherwise overwrite each other
      - skips files already present, so the export can be re-run incrementally
      - optionally moves the processed messages to another folder

    Outlook COM automation requires the script and Outlook to run at the SAME
    Windows integrity level. A PowerShell session started with "Run as
    administrator" cannot attach to a normally started Outlook and fails with
    CO_E_SERVER_EXEC_FAILURE (0x80080005). The script detects this up front.

.PARAMETER FolderPath
    Folder to export, given as a backslash-separated path from the mailbox root,
    for example "dmarc_rua@contoso.com\Inbox\reports". When omitted, Outlook
    prompts for the folder.

.PARAMETER OutputPath
    Destination directory. Created when missing. Defaults to .\DmarcReports

.PARAMETER Since
    Only export messages received on or after this date.

.PARAMETER Until
    Only export messages received before this date.

.PARAMETER Extensions
    Attachment extensions to export. Defaults to gz, zip and xml.

.PARAMETER MoveToFolder
    Folder path the exported messages are moved to, relative to the same
    mailbox root. The folder must already exist.

.PARAMETER IncludeSubfolders
    Also processes the subfolders of the selected folder.

.PARAMETER Overwrite
    Re-exports attachments even when a file of the same name already exists.

.EXAMPLE
    .\Export-DmarcAttachments.ps1

.EXAMPLE
    .\Export-DmarcAttachments.ps1 -FolderPath "dmarc_rua@contoso.com\Inbox" -OutputPath C:\DMARC\raw

.EXAMPLE
    .\Export-DmarcAttachments.ps1 -FolderPath "dmarc_rua@contoso.com\Inbox" -Since 2026-08-01 -MoveToFolder "dmarc_rua@contoso.com\Inbox\processed"

.NOTES
    Version 1.0.0
    Requires the Outlook desktop client (classic). The new Outlook for Windows
    does not expose COM automation; use Export-DmarcAttachmentsGraph.ps1 or an
    Outlook VBA macro in that case.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]   $FolderPath,
    [string]   $OutputPath = ".\DmarcReports",
    [datetime] $Since,
    [datetime] $Until,
    [string[]] $Extensions = @("gz", "zip", "xml"),
    [string]   $MoveToFolder,
    [switch]   $IncludeSubfolders,
    [switch]   $Overwrite
)

Set-StrictMode -Off
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Pre-flight checks
# ---------------------------------------------------------------------------

# An elevated session cannot drive a normally started Outlook.
$identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw @"
This PowerShell session is running elevated.

Outlook automation fails from an elevated session (CO_E_SERVER_EXEC_FAILURE,
0x80080005) because Windows blocks the COM connection between processes at
different integrity levels.

Close this window and start PowerShell normally, without 'Run as administrator'.
"@
}

if (-not (Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue)) {
    Write-Warning "Outlook does not appear to be running. Start it and sign in before continuing."
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Get-OutlookFolderByPath {
    <#
        Walks a backslash-separated path from the mailbox root.
        The first segment is the store display name (usually the mailbox
        address); the remaining segments are folders.
    #>
    param(
        [Parameter(Mandatory)] $Namespace,
        [Parameter(Mandatory)] [string] $Path
    )

    $segments = $Path.Trim('\') -split '\\'
    $storeName = $segments[0]

    $root = $null
    foreach ($f in $Namespace.Folders) {
        if ($f.Name -eq $storeName) { $root = $f; break }
    }
    if ($null -eq $root) {
        $available = ($Namespace.Folders | ForEach-Object { $_.Name }) -join ", "
        throw "Mailbox '$storeName' not found. Available: $available"
    }

    $current = $root
    foreach ($segment in $segments[1..($segments.Count - 1)]) {
        $next = $null
        foreach ($sub in $current.Folders) {
            if ($sub.Name -eq $segment) { $next = $sub; break }
        }
        if ($null -eq $next) {
            $available = ($current.Folders | ForEach-Object { $_.Name }) -join ", "
            throw "Folder '$segment' not found under '$($current.Name)'. Available: $available"
        }
        $current = $next
    }
    return $current
}

function Get-SafeFileName {
    param([string] $Name)
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $Name.ToCharArray()) {
        if ($invalid -contains $c) { [void]$sb.Append('_') } else { [void]$sb.Append($c) }
    }
    $clean = $sb.ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($clean)) { $clean = "attachment" }
    return $clean
}

# ---------------------------------------------------------------------------
# Connect to Outlook
# ---------------------------------------------------------------------------

try {
    $outlook = New-Object -ComObject Outlook.Application
}
catch {
    throw @"
Unable to connect to Outlook: $($_.Exception.Message)

Checklist:
  - Outlook classic must be installed and running, with a configured profile
  - this session must NOT be elevated
  - no modal dialog (password prompt, Safe Mode) may be open in Outlook
  - PowerShell and Outlook must share the same architecture (both 64-bit)
  - the new Outlook for Windows does not support COM automation
"@
}

$namespace = $outlook.GetNamespace("MAPI")

if ($FolderPath) {
    $folder = Get-OutlookFolderByPath -Namespace $namespace -Path $FolderPath
}
else {
    Write-Host "Select the folder holding the DMARC reports..." -ForegroundColor Cyan
    $folder = $namespace.PickFolder()
    if ($null -eq $folder) { Write-Host "Cancelled."; return }
}

$targetFolder = $null
if ($MoveToFolder) {
    $targetFolder = Get-OutlookFolderByPath -Namespace $namespace -Path $MoveToFolder
}

# ---------------------------------------------------------------------------
# Prepare output
# ---------------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}
$OutputPath = (Resolve-Path -LiteralPath $OutputPath).Path

$wanted = $Extensions | ForEach-Object { "." + $_.TrimStart('.').ToLower() }

$stats = [ordered]@{
    Messages    = 0
    Attachments = 0
    Exported    = 0
    Skipped     = 0
    Failed      = 0
    Moved       = 0
}
$errors = New-Object System.Collections.Generic.List[string]

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

function Export-Folder {
    param($Folder)

    Write-Host "Folder: $($Folder.FolderPath)" -ForegroundColor Cyan

    $items = $Folder.Items

    # Restrict by date in MAPI rather than filtering in PowerShell: on a RUA
    # mailbox with tens of thousands of items this is the difference between
    # seconds and many minutes.
    $clauses = @()
    if ($Since) { $clauses += "[ReceivedTime] >= '" + $Since.ToString("g") + "'" }
    if ($Until) { $clauses += "[ReceivedTime] < '"  + $Until.ToString("g") + "'" }
    if ($clauses.Count -gt 0) {
        $filter = $clauses -join " AND "
        Write-Verbose "Restrict: $filter"
        $items = $items.Restrict($filter)
    }

    $total = $items.Count
    Write-Host "  messages in scope: $total"

    # Iterate backwards: moving items mutates the collection, and going from the
    # end keeps the remaining indexes valid.
    for ($i = $total; $i -ge 1; $i--) {

        try { $item = $items.Item($i) } catch { continue }
        if ($null -eq $item) { continue }

        # only mail items carry attachments we care about
        if ($item.Class -ne 43) { continue }

        $stats.Messages++

        if (($total - $i + 1) % 250 -eq 0) {
            Write-Host "  processed $($total - $i + 1) / $total..." -ForegroundColor DarkGray
        }

        $received = $null
        try { $received = $item.ReceivedTime } catch { }
        $stamp = if ($received) { $received.ToString("yyyyMMdd-HHmmss") } else { "nodate" }

        $savedAny = $false

        try   { $attachments = $item.Attachments }
        catch { continue }

        for ($a = 1; $a -le $attachments.Count; $a++) {

            $att = $attachments.Item($a)
            $name = $att.FileName
            if ([string]::IsNullOrWhiteSpace($name)) { continue }

            $ext = [System.IO.Path]::GetExtension($name).ToLower()
            if ($wanted -notcontains $ext) { continue }

            $stats.Attachments++

            $fileName = Get-SafeFileName -Name ("{0}-{1}" -f $stamp, $name)
            $destination = Join-Path $OutputPath $fileName

            if ((Test-Path -LiteralPath $destination) -and (-not $Overwrite)) {
                $stats.Skipped++
                $savedAny = $true
                continue
            }

            # two messages in the same second with the same attachment name
            if ((Test-Path -LiteralPath $destination) -and $Overwrite) {
                $base = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
                $n = 1
                while (Test-Path -LiteralPath $destination) {
                    $destination = Join-Path $OutputPath ("{0}({1}){2}" -f $base, $n, $ext)
                    $n++
                }
            }

            if ($PSCmdlet.ShouldProcess($destination, "Save attachment")) {
                try {
                    $att.SaveAsFile($destination)
                    $stats.Exported++
                    $savedAny = $true
                }
                catch {
                    $stats.Failed++
                    $errors.Add("$fileName -> $($_.Exception.Message)")
                }
            }
        }

        if ($savedAny -and $targetFolder) {
            if ($PSCmdlet.ShouldProcess($item.Subject, "Move message")) {
                try   { [void]$item.Move($targetFolder); $stats.Moved++ }
                catch { $errors.Add("move '$($item.Subject)' -> $($_.Exception.Message)") }
            }
        }
    }

    if ($IncludeSubfolders) {
        foreach ($sub in $Folder.Folders) { Export-Folder -Folder $sub }
    }
}

Export-Folder -Folder $folder

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "=== SUMMARY ===" -ForegroundColor Cyan
Write-Host ("Messages examined   : {0}" -f $stats.Messages)
Write-Host ("Attachments matched : {0}" -f $stats.Attachments)
Write-Host ("Exported            : {0}" -f $stats.Exported) -ForegroundColor Green
Write-Host ("Already present     : {0}" -f $stats.Skipped)
if ($targetFolder) { Write-Host ("Messages moved      : {0}" -f $stats.Moved) }
Write-Host ("Failed              : {0}" -f $stats.Failed) -ForegroundColor $(if ($stats.Failed -gt 0) { "Yellow" } else { "Green" })
Write-Host ""
Write-Host "Output: $OutputPath"

if ($errors.Count -gt 0) {
    Write-Host ""
    Write-Warning "Errors: $($errors.Count)"
    $errors | Select-Object -First 20 | ForEach-Object { Write-Host "  - $_" -ForegroundColor DarkYellow }
    if ($errors.Count -gt 20) { Write-Host "  ... and $($errors.Count - 20) more" -ForegroundColor DarkYellow }
}

if ($stats.Exported -gt 0) {
    Write-Host ""
    Write-Host "Next step:" -ForegroundColor Cyan
    Write-Host "  ..\Analyze-DmarcReports\Analyze-DmarcReports.ps1 -Path `"$OutputPath`"" -ForegroundColor Gray
}

# release the COM references
foreach ($ref in @($targetFolder, $folder, $namespace, $outlook)) {
    if ($ref) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ref) }
}
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
