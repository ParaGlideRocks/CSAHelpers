#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Automates migration batches to migrate mailboxes from Exchange On-Premises to Office 365.

.DESCRIPTION
    Reads a text file containing user email addresses (one per line), creates migration
    batches in Exchange Online, monitors progress, and provides comprehensive logging.
    Supports batching users into configurable group sizes to avoid overloading the
    migration service.

.PARAMETER UserFile
    Path to the text file containing one user email address (UPN) per line.

.PARAMETER MigrationEndpoint
    The Migration Endpoint identity (Name) that was created in Exchange Online
    using New-MigrationEndpoint (for example, "OnpremEndpoint"). This value is
    passed to New-MigrationBatch as -SourceEndpoint.

.PARAMETER TargetDeliveryDomain
    The target delivery domain for the migration (e.g., "contoso.mail.onmicrosoft.com").

.PARAMETER BatchNamePrefix
    Prefix for naming migration batches. Defaults to "MigBatch".

.PARAMETER BatchSize
    Number of users per migration batch. Defaults to 50.

.PARAMETER NotificationEmails
    Email addresses to receive migration batch notifications.

.PARAMETER AutoStart
    Automatically start the migration batch after creation.

.PARAMETER AutoComplete
    Automatically complete the migration batch after the initial synchronization finishes.

.PARAMETER LogFile
    Path to the log file. Defaults to a timestamped file in the script directory.

.PARAMETER BadItemLimit
    (Deprecated in Exchange Online) Maximum number of bad items to skip per mailbox.
    If omitted, the script won't pass BadItemLimit to New-MigrationBatch.

.PARAMETER LargeItemLimit
    (Deprecated in Exchange Online) Maximum number of large items to skip per mailbox.
    If omitted, the script won't pass LargeItemLimit to New-MigrationBatch.

.PARAMETER WhatIf
    Preview changes without actually creating migration batches.

.EXAMPLE
    .\Start-EXOMigrationBatch.ps1 -UserFile "C:\Migration\Users.txt" `
        -MigrationEndpoint "OnpremEndpoint" `
        -TargetDeliveryDomain "contoso.mail.onmicrosoft.com"

.EXAMPLE
    .\Start-EXOMigrationBatch.ps1 -UserFile "C:\Migration\Users.txt" `
        -MigrationEndpoint "OnpremEndpoint" `
        -TargetDeliveryDomain "contoso.mail.onmicrosoft.com" `
        -BatchSize 25 -AutoStart -AutoComplete `
        -NotificationEmails "admin@contoso.com"

.EXAMPLE
    .\Start-EXOMigrationBatch.ps1 -UserFile "C:\Migration\Users.txt" `
        -MigrationEndpoint "OnpremEndpoint" `
        -TargetDeliveryDomain "contoso.mail.onmicrosoft.com" -WhatIf

.NOTES
    Author:  EXOHelpers
    Requires: Exchange Online PowerShell V3 module
    Permissions: Exchange Administrator or Global Administrator
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Path to text file with user UPNs, one per line.")]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$UserFile,

    [Parameter(Mandatory = $true, HelpMessage = "Migration endpoint identity (Name) in Exchange Online (for example, 'OnpremEndpoint').")]
    [string]$MigrationEndpoint,

    [Parameter(Mandatory = $true, HelpMessage = "Target delivery domain (e.g. contoso.mail.onmicrosoft.com).")]
    [string]$TargetDeliveryDomain,

    [Parameter(Mandatory = $false)]
    [string]$BatchNamePrefix = "MigBatch",

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 200)]
    [int]$BatchSize = 50,

    [Parameter(Mandatory = $false)]
    [string[]]$NotificationEmails,

    [Parameter(Mandatory = $false)]
    [switch]$AutoStart,

    [Parameter(Mandatory = $false)]
    [switch]$AutoComplete,

    [Parameter(Mandatory = $false)]
    [string]$LogFile = (Join-Path $PSScriptRoot ("Start-EXOMigrationBatch_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))),

    # NOTE: BadItemLimit and LargeItemLimit are deprecated in Exchange Online.
    # Microsoft recommends reviewing the Data Consistency Score instead.
    # These parameters are kept for backward compatibility but may be removed in the future.
    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 1000)]
    [int]$BadItemLimit,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 1000)]
    [int]$LargeItemLimit
)

#region ── Logging ──

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped, leveled message to both console and log file.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry     = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        "ERROR"   { Write-Host $entry -ForegroundColor Red }
        "WARNING" { Write-Host $entry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $entry -ForegroundColor Green }
        default   { Write-Host $entry -ForegroundColor Cyan }
    }

    try {
        $logDir = Split-Path -Path $LogFile -Parent
        if ($logDir -and -not (Test-Path -Path $logDir -PathType Container)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }

        Add-Content -Path $LogFile -Value $entry -ErrorAction Stop
    }
    catch {
        Write-Host "[WARNING] Unable to write to log file '$LogFile': $_" -ForegroundColor Yellow
    }
}

#endregion

#region ── Input Helpers ──

function Read-UserList {
    <#
    .SYNOPSIS
        Reads and validates the user list from the input text file.
    .OUTPUTS
        Array of trimmed, non-empty, unique email addresses.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    Write-Log "Reading user list from '$Path'..."

    try {
        $raw = Get-Content -Path $Path -ErrorAction Stop

        # Remove blank lines, trim whitespace, and deduplicate
        $users = $raw |
            ForEach-Object { $_.Trim() } |
            Where-Object  { $_ -ne "" }  |
            Select-Object -Unique

        if (-not $users -or $users.Count -eq 0) {
            Write-Log "User file is empty or contains no valid entries." -Level ERROR
            return $null
        }

        Write-Log "Found $($users.Count) unique user(s) in file." -Level SUCCESS
        return @($users)
    }
    catch {
        Write-Log "Failed to read user file: $_" -Level ERROR
        return $null
    }
}

function Test-UserValid {
    <#
    .SYNOPSIS
        Validates that a string looks like an email / UPN.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Identity
    )

    return $Identity -match '^[^@\s]+@[^@\s]+\.[^@\s]+$'
}

#endregion

#region ── Batch Helpers ──

function Split-IntoBatches {
    <#
    .SYNOPSIS
        Splits an array into chunks of a given size.
    #>
    param(
        [Parameter(Mandatory)]
        [object[]]$Items,

        [Parameter(Mandatory)]
        [ValidateRange(1, [int]::MaxValue)]
        [int]$Size
    )

    $batches = [System.Collections.Generic.List[object]]::new()

    if (-not $Items -or $Items.Count -eq 0) {
        return ,$batches
    }

    for ($i = 0; $i -lt $Items.Count; $i += $Size) {
        $end   = [Math]::Min($i + $Size, $Items.Count) - 1
        $chunk = @($Items[$i..$end])
        $null = $batches.Add($chunk)
    }

    # Return as a single object so PowerShell doesn't enumerate/flatten it
    return ,$batches
}

function New-MigrationCSV {
    <#
    .SYNOPSIS
        Creates a temporary CSV file formatted for New-MigrationBatch.
    .OUTPUTS
        Full path to the generated CSV.
    #>
    param(
        [Parameter(Mandatory)]
        [string[]]$Users,

        [Parameter(Mandatory)]
        [string]$BatchName
    )

    $csvPath = Join-Path ([System.IO.Path]::GetTempPath()) "$BatchName.csv"

    try {
        $Users | ForEach-Object {
            [PSCustomObject]@{ EmailAddress = $_ }
        } | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop

        Write-Log "CSV file created: $csvPath"
        return $csvPath
    }
    catch {
        Write-Log "Failed to create CSV for batch '$BatchName': $_" -Level ERROR
        return $null
    }
}

#endregion

#region ── Migration Endpoint ──

function Resolve-MigrationEndpoint {
    <#
    .SYNOPSIS
        Verifies the migration endpoint exists.
    .OUTPUTS
        The MigrationEndpoint object, or $null on failure.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$EndpointName
    )

    Write-Log "Resolving migration endpoint '$EndpointName'..."

    try {
        $endpoint = Get-MigrationEndpoint -Identity $EndpointName -ErrorAction Stop
        Write-Log "Migration endpoint '$EndpointName' found (Type: $($endpoint.EndpointType))." -Level SUCCESS
        return $endpoint
    }
    catch {
        Write-Log "Migration endpoint '$EndpointName' not found. Please create it first using New-MigrationEndpoint." -Level ERROR
        Write-Log "Example:  $cred = Get-Credential; New-MigrationEndpoint -ExchangeRemoteMove -Name 'OnpremEndpoint' -Autodiscover -EmailAddress admin@contoso.com -Credentials $cred" -Level WARNING
        return $null
    }
}

#endregion

#region ── Batch Creation ──

function New-EXOMigrationBatch {
    <#
    .SYNOPSIS
        Creates a single migration batch in Exchange Online.
    .OUTPUTS
        $true on success, $false on failure.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)]
        [string]$BatchName,

        [Parameter(Mandatory)]
        [string]$CSVPath,

        [Parameter(Mandatory)]
        [string]$TargetDelivery,

        [Parameter(Mandatory)]
        [string]$Endpoint,

        [int]$BadItems,

        [int]$LargeItems,

        [string[]]$Notifications,

        [switch]$Start,

        [switch]$Complete
    )

    if (-not (Test-Path -Path $CSVPath -PathType Leaf)) {
        Write-Log "CSVPath not found for batch '$BatchName': $CSVPath" -Level ERROR
        return $false
    }

    try {
        $csvData = [System.IO.File]::ReadAllBytes($CSVPath)

        $params = @{
            Name                 = $BatchName
            SourceEndpoint       = $Endpoint
            TargetDeliveryDomain = $TargetDelivery
            CSVData              = $csvData
            ErrorAction          = "Stop"
        }

        # BadItemLimit and LargeItemLimit are deprecated in Exchange Online.
        # Only include them if explicitly provided by the caller.
        if ($PSBoundParameters.ContainsKey('BadItems')) {
            Write-Log "BadItemLimit is deprecated in Exchange Online. Consider using Data Consistency Score instead." -Level WARNING
            $params["BadItemLimit"] = $BadItems
        }

        if ($PSBoundParameters.ContainsKey('LargeItems')) {
            Write-Log "LargeItemLimit is deprecated in Exchange Online. Consider using Data Consistency Score instead." -Level WARNING
            $params["LargeItemLimit"] = $LargeItems
        }

        if ($Notifications) {
            $params["NotificationEmails"] = $Notifications
        }

        if ($Start) {
            $params["AutoStart"] = $true
        }

        if ($Complete) {
            $params["AutoComplete"] = $true
        }

        Write-Log "Creating migration batch '$BatchName' ($Endpoint -> $TargetDelivery)..."

        New-MigrationBatch @params | Out-Null

        if ($WhatIfPreference) {
            Write-Log "[WhatIf] Would create migration batch '$BatchName'." -Level WARNING
        }
        else {
            Write-Log "Migration batch '$BatchName' created successfully." -Level SUCCESS
        }

        return $true
    }
    catch {
        Write-Log "Failed to create batch '$BatchName': $_" -Level ERROR
        return $false
    }
}

#endregion

#region ── Status Reporting ──

function Get-BatchStatus {
    <#
    .SYNOPSIS
        Retrieves and logs the current status of all batches created in this run.
    #>
    param(
        [Parameter(Mandatory)]
        [string[]]$BatchNames
    )

    Write-Log "── Migration Batch Status Report ──"

    $report = foreach ($name in $BatchNames) {
        try {
            $batch = Get-MigrationBatch -Identity $name -ErrorAction Stop

            [PSCustomObject]@{
                BatchName    = $batch.Identity
                Status       = $batch.Status
                TotalCount   = $batch.TotalCount
                SyncedCount  = $batch.SyncedItemCount
                FailedCount  = $batch.FailedItemCount
                State        = $batch.State
            }
        }
        catch {
            Write-Log "Could not retrieve status for batch '$name': $_" -Level WARNING

            [PSCustomObject]@{
                BatchName    = $name
                Status       = "Unknown"
                TotalCount   = "N/A"
                SyncedCount  = "N/A"
                FailedCount  = "N/A"
                State        = "Error retrieving"
            }
        }
    }

    $report | Format-Table -AutoSize | Out-String | ForEach-Object { Write-Log $_ }

    # Export status to CSV alongside the log
    $statusFile = $LogFile -replace '\.log$', '_Status.csv'
    $report | Export-Csv -Path $statusFile -NoTypeInformation -Encoding UTF8 -ErrorAction SilentlyContinue
    Write-Log "Status report exported to '$statusFile'."

    return $report
}

#endregion

#region ── Connection Validation ──

function Assert-ExchangeOnlineConnection {
    <#
    .SYNOPSIS
        Verifies that an active Exchange Online session exists.
    #>
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Log "Exchange Online session is active." -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "No active Exchange Online session detected. Please run Connect-ExchangeOnline first." -Level ERROR
        return $false
    }
}

#endregion

#region ── Main Execution ──

try {
    # ── Header ──
    Write-Log "═══════════════════════════════════════════════════════════"
    Write-Log "  Exchange On-Premises → Office 365 Migration Automation  "
    Write-Log "═══════════════════════════════════════════════════════════"
    Write-Log "Script started at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Log "User file      : $UserFile"
    Write-Log "Source endpoint: $MigrationEndpoint"
    Write-Log "Target domain  : $TargetDeliveryDomain"
    Write-Log "Batch size     : $BatchSize"
    if ($PSBoundParameters.ContainsKey('BadItemLimit'))   { Write-Log "Bad item limit : $BadItemLimit (deprecated in EXO)" -Level WARNING }
    if ($PSBoundParameters.ContainsKey('LargeItemLimit')) { Write-Log "Large item lim : $LargeItemLimit (deprecated in EXO)" -Level WARNING }
    Write-Log "Auto-start     : $AutoStart"
    Write-Log "Auto-complete  : $AutoComplete"
    Write-Log "Log file       : $LogFile"

    # ── Step 1 – Verify EXO connection ──
    if (-not (Assert-ExchangeOnlineConnection)) {
        throw "Exchange Online connection required. Run Connect-ExchangeOnline and retry."
    }

    # ── Step 2 – Resolve migration endpoint ──
    $endpoint = Resolve-MigrationEndpoint -EndpointName $MigrationEndpoint
    if (-not $endpoint) {
        throw "Migration endpoint '$MigrationEndpoint' could not be resolved."
    }

    # ── Step 3 – Read & validate user list ──
    $users = Read-UserList -Path $UserFile
    if (-not $users) {
        throw "No valid users found in '$UserFile'."
    }

    # Validate email format
    $validUsers   = [System.Collections.Generic.List[string]]::new()
    $invalidUsers = [System.Collections.Generic.List[string]]::new()

    foreach ($user in $users) {
        if (Test-UserValid -Identity $user) {
            $validUsers.Add($user)
        }
        else {
            $invalidUsers.Add($user)
            Write-Log "Invalid email format skipped: '$user'" -Level WARNING
        }
    }

    if ($invalidUsers.Count -gt 0) {
        Write-Log "$($invalidUsers.Count) user(s) skipped due to invalid email format." -Level WARNING
    }

    if ($validUsers.Count -eq 0) {
        throw "No valid email addresses remaining after validation."
    }

    Write-Log "$($validUsers.Count) valid user(s) will be processed." -Level SUCCESS

    # ── Step 4 – Split into batches ──
    $batches = Split-IntoBatches -Items $validUsers.ToArray() -Size $BatchSize
    Write-Log "Users split into $($batches.Count) batch(es) of up to $BatchSize."

    # ── Step 5 – Create migration batches ──
    $timestamp       = Get-Date -Format "yyyyMMdd_HHmmss"
    $createdBatches  = [System.Collections.Generic.List[string]]::new()
    $failedBatches   = [System.Collections.Generic.List[string]]::new()

    for ($b = 0; $b -lt $batches.Count; $b++) {
        $batchNumber = $b + 1
        $batchName   = "{0}_{1}_{2}" -f $BatchNamePrefix, $timestamp, $batchNumber
        $batchUsers  = $batches[$b]

        Write-Log "── Batch $batchNumber / $($batches.Count): '$batchName' ($($batchUsers.Count) users) ──"

        # Create CSV for this batch
        $csvPath = New-MigrationCSV -Users $batchUsers -BatchName $batchName
        if (-not $csvPath) {
            Write-Log "Skipping batch '$batchName' due to CSV creation failure." -Level ERROR
            $failedBatches.Add($batchName)
            continue
        }

        # Create the migration batch
        # Build optional splat for deprecated parameters
        $optionalParams = @{}
        if ($PSBoundParameters.ContainsKey('BadItemLimit'))   { $optionalParams['BadItems']  = $BadItemLimit }
        if ($PSBoundParameters.ContainsKey('LargeItemLimit')) { $optionalParams['LargeItems'] = $LargeItemLimit }

        $success = New-EXOMigrationBatch `
            -BatchName      $batchName `
            -CSVPath        $csvPath `
            -TargetDelivery $TargetDeliveryDomain `
            -Endpoint       $MigrationEndpoint `
            -Notifications  $NotificationEmails `
            -Start:$AutoStart `
            -Complete:$AutoComplete `
            @optionalParams

        if ($success) {
            $createdBatches.Add($batchName)
        }
        else {
            $failedBatches.Add($batchName)
        }

        # Clean up temp CSV
        Remove-Item -Path $csvPath -Force -ErrorAction SilentlyContinue
    }

    # ── Step 6 – Summary ──
    Write-Log "═══════════════════════════════════════════════════════════"
    Write-Log "  Migration Batch Creation Summary"
    Write-Log "═══════════════════════════════════════════════════════════"
    Write-Log "Total batches attempted : $($batches.Count)"
    Write-Log "Successfully created    : $($createdBatches.Count)" -Level SUCCESS
    Write-Log "Failed                  : $($failedBatches.Count)"  -Level $(if ($failedBatches.Count -gt 0) { "ERROR" } else { "INFO" })

    if ($failedBatches.Count -gt 0) {
        Write-Log "Failed batches: $($failedBatches -join ', ')" -Level ERROR
    }

    # ── Step 7 – Status report (for successfully created batches) ──
    if ($createdBatches.Count -gt 0 -and -not $WhatIfPreference) {
        Start-Sleep -Seconds 5   # brief pause to let EXO register batches
        Get-BatchStatus -BatchNames $createdBatches.ToArray()
    }

    Write-Log "Script completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')." -Level SUCCESS
}
catch {
    Write-Log "FATAL: $($_.Exception.Message)" -Level ERROR
    Write-Log "Stack: $($_.ScriptStackTrace)" -Level ERROR
    exit 1
}

#endregion
