<#
.SYNOPSIS
    Removes SMTP addresses associated with invalid domains from Exchange users.

.DESCRIPTION
    Reads a list of UPNs from a text file, retrieves their SMTP proxy addresses,
    and removes any addresses matching the specified invalid domains.

.NOTES
        DISCLAIMER
        This script is provided "as-is" without warranty of any kind. You are solely responsible
        for reviewing, testing, and validating results before and after running it in your tenant/
        environment.

        - Always run with `-WhatIf` first (supported via ShouldProcess) to review intended changes.
        - This script can modify mailbox configuration (proxy addresses) and may impact mail flow.
        - If a mailbox has Email Address Policy enabled, the script disables it and does not
            re-enable it automatically.
        - Use least-privilege admin roles and run in a non-production/test scope when possible.

.PARAMETER UserFile
    Path to the text file containing one UPN per line.

.PARAMETER InvalidDomains
    Array of invalid domain names (e.g., "olddomain.com", "deprecated.org").

.PARAMETER LogFile
    Path to the log file. Defaults to a timestamped file in the script directory.

.EXAMPLE
    .\Remove-InvalidSMTP.ps1 -UserFile "C:\Users.txt" -InvalidDomains "olddomain.com","legacy.org"

#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [string]$UserFile,

    [Parameter(Mandatory = $true)]
    [string[]]$InvalidDomains,

    [Parameter(Mandatory = $false)]
    [string]$LogFile = (Join-Path $PSScriptRoot ("Remove-InvalidSMTP_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss")))
)

#region Logging Functions

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        "ERROR"   { Write-Host $entry -ForegroundColor Red }
        "WARNING" { Write-Host $entry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $entry -ForegroundColor Green }
        default   { Write-Host $entry -ForegroundColor Cyan }
    }

    Add-Content -Path $LogFile -Value $entry -ErrorAction SilentlyContinue
}

function Write-LogSeparator {
    $separator = "=" * 80
    Add-Content -Path $LogFile -Value $separator -ErrorAction SilentlyContinue
    Write-Host $separator -ForegroundColor DarkGray
}

#endregion
#region Main Execution

Write-LogSeparator
Write-Log "Script execution started."
Write-Log "User file       : $UserFile"
Write-Log "Invalid domains : $($InvalidDomains -join ', ')"
Write-Log "Log file        : $LogFile"
Write-LogSeparator

# --- Validate user file ---
if (-not (Test-Path -Path $UserFile -PathType Leaf)) {
    Write-Log "User file not found: '$UserFile'. Aborting." -Level ERROR
    exit 1
}

$upnList = Get-Content -Path $UserFile | Where-Object { $_.Trim() -ne "" }

if ($upnList.Count -eq 0) {
    Write-Log "User file is empty. No users to process. Aborting." -Level ERROR
    exit 1
}

Write-Log "Loaded $($upnList.Count) UPN(s) from file."

# --- Validate Exchange connectivity ---
try {
    $null = Get-Command Get-Mailbox -ErrorAction Stop
    Write-Log "Exchange cmdlets are available." -Level SUCCESS
}
catch {
    Write-Log "Exchange cmdlets not found. Ensure you are connected to Exchange (on-prem or Exchange Online). Aborting." -Level ERROR
    Write-Log "Error: $($_.Exception.Message)" -Level ERROR
    exit 1
}

# --- Counters ---
$totalProcessed    = 0
$totalRemoved      = 0
$totalSkipped      = 0
$totalUserNotFound = 0
$totalErrors       = 0

# --- Process each UPN ---
foreach ($upn in $upnList) {
    $upn = $upn.Trim()
    $totalProcessed++
    Write-LogSeparator
    Write-Log "Processing user ($totalProcessed/$($upnList.Count)): $upn"

    # Check if the mailbox exists
    try {
        $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop
    }
    catch {
        Write-Log "User '$upn' not found or has no mailbox. Skipping." -Level WARNING
        Write-Log "Error detail: $($_.Exception.Message)" -Level WARNING
        $totalUserNotFound++
        continue
    }

    # Retrieve current SMTP proxy addresses
    $proxyAddresses = $mailbox.EmailAddresses | Where-Object { $_ -like "smtp:*" -or $_ -like "SMTP:*" }

    if (-not $proxyAddresses -or $proxyAddresses.Count -eq 0) {
        Write-Log "No SMTP addresses found for '$upn'. Skipping." -Level WARNING
        $totalSkipped++
        continue
    }

    Write-Log "Current SMTP addresses for '$upn':"
    foreach ($addr in $proxyAddresses) {
        Write-Log "  $addr"
    }

    # Identify addresses to remove (matching invalid domains)
    $addressesToRemove = @()
    foreach ($addr in $proxyAddresses) {
        $emailPart = ($addr -replace "^smtp:", "" -replace "^SMTP:", "").Trim()
        $domain = ($emailPart -split "@")[-1]

        foreach ($invalidDomain in $InvalidDomains) {
            if ($domain -ieq $invalidDomain) {
                # Prevent removal of the primary SMTP address
                if ($addr -cmatch "^SMTP:") {
                    Write-Log "  SKIPPING primary address '$addr' — cannot remove the primary SMTP address." -Level WARNING
                    $totalSkipped++
                }
                else {
                    $addressesToRemove += $addr
                }
                break
            }
        }
    }

    if ($addressesToRemove.Count -eq 0) {
        Write-Log "No invalid-domain addresses to remove for '$upn'." -Level INFO
        continue
    }

    # --- Disable Email Address Policy before modifying addresses ---
    if ($mailbox.EmailAddressPolicyEnabled -eq $true) {
        if ($PSCmdlet.ShouldProcess($upn, "Disable EmailAddressPolicy")) {
            try {
                Set-Mailbox -Identity $upn -EmailAddressPolicyEnabled $false -ErrorAction Stop
                Write-Log "Disabled Email Address Policy for '$upn'. It will NOT be re-enabled." -Level SUCCESS
            }
            catch {
                Write-Log "Failed to disable Email Address Policy for '$upn': $($_.Exception.Message)" -Level ERROR
                Write-Log "Skipping address removal for this user to avoid policy conflicts." -Level ERROR
                $totalErrors++
                continue
            }
        }
    }
    else {
        Write-Log "Email Address Policy already disabled for '$upn'." -Level INFO
    }

    # --- Remove each invalid address ---
    foreach ($addrToRemove in $addressesToRemove) {
        Write-Log "  Removing: $addrToRemove"

        if ($PSCmdlet.ShouldProcess($upn, "Remove address $addrToRemove")) {
            try {
                Set-Mailbox -Identity $upn -EmailAddresses @{Remove = $addrToRemove} -ErrorAction Stop
                Write-Log "  Removed successfully: $addrToRemove" -Level SUCCESS
                $totalRemoved++
            }
            catch {
                Write-Log "  Failed to remove '$addrToRemove' from '$upn': $($_.Exception.Message)" -Level ERROR
                $totalErrors++
            }
        }
        else {
            Write-Log "  [WhatIf] Would remove: $addrToRemove" -Level WARNING
        }
    }
}

#endregion

#region Summary

Write-LogSeparator
Write-Log "Script execution completed."
Write-Log "Summary:"
Write-Log "  Total users processed  : $totalProcessed"
Write-Log "  Users not found        : $totalUserNotFound"
Write-Log "  Addresses removed      : $totalRemoved"
Write-Log "  Addresses skipped      : $totalSkipped"
Write-Log "  Errors                 : $totalErrors"
Write-LogSeparator

#endregion
