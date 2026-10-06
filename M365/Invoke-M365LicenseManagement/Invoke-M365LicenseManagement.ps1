#requires -Version 7.0

<#
.SYNOPSIS
    Interactive Microsoft 365 license management through Microsoft Graph.

.DESCRIPTION
    Menu:
      1. Show the licenses (SKUs) available in the tenant
      2. Export the users that hold a selected license
      3. Replace a license with another one, keeping every other license
      4. Restore licenses from a replacement backup (rollback)
      5. Exit

    License replacement:
      - adds the target SKU and removes ONLY the source SKU, in a single call;
      - never touches any other SKU assigned to the user;
      - skips users whose source license is (also) inherited from a group;
      - skips users without UsageLocation;
      - skips disabled accounts unless -IncludeDisabledAccounts is used;
      - can be scoped to a pilot (-PilotUserCsv and/or -PilotGroupId);
      - by default carries the source "disabled service plans" over to the
        target SKU when the same service plan exists in the target
        (-DisabledPlansMode Preserve | None);
      - reports, per user, the service plans that would be LOST
        (enabled in the source, not provided by the target or by any
        other SKU still assigned to the user);
      - always produces a pre-check CSV and a JSON backup before any change;
      - supports -WhatIf for a formal dry-run without writes;
      - makes no change until the operator types EXECUTE;
      - produces a result CSV and a log file.

.PARAMETER OutputFolder
    Folder for CSV reports, JSON backups and logs.

.PARAMETER PilotUserCsv
    Optional CSV file with a "UserPrincipalName" column (";" or "," delimited).
    When provided, only these users are considered for the replacement.

.PARAMETER PilotGroupId
    Optional Entra ID group object ID. When provided, only the (transitive)
    members of this group are considered for the replacement.
    If both -PilotUserCsv and -PilotGroupId are provided, a user must match
    at least one of the two.

.PARAMETER DisabledPlansMode
    Preserve (default): service plans disabled on the source SKU are disabled
                        on the target SKU too, when they exist in the target.
    None:               the target SKU is assigned with all service plans enabled.

.PARAMETER IncludeDisabledAccounts
    Also process users whose account is disabled (they will consume a target license).

.PARAMETER IncludeInactiveSkus
    Also list SKUs whose CapabilityStatus is not Enabled/Warning (e.g. Suspended).

.PARAMETER MaxRetryCount
    Number of retries for transient Microsoft Graph errors such as throttling
    and temporary service failures.

.PARAMETER RetryBaseDelaySeconds
    Base delay used for exponential backoff between Graph retries.

.PARAMETER RequestDelayMilliseconds
    Delay after each write operation to reduce Graph request pressure.

.PARAMETER BatchSize
    Number of write operations to perform before pausing. Use 0 to disable
    batch pauses.

.PARAMETER BatchPauseSeconds
    Pause after each completed batch of write operations.

.NOTES
    Microsoft Graph delegated permissions (least privilege):
      User.Read.All                    - read users and their licenses
      Organization.Read.All            - read subscribed SKUs
      LicenseAssignment.ReadWrite.All  - assign / remove licenses
      GroupMember.Read.All             - only when -PilotGroupId is used

    Required Entra ID role for the signed-in account: License Administrator
    (or User Administrator).

    Always run first against a pilot group.

.EXAMPLE
    ./Invoke-M365LicenseManagement.ps1 -PilotUserCsv .\pilot.csv

.EXAMPLE
    ./Invoke-M365LicenseManagement.ps1 -PilotGroupId 00000000-0000-0000-0000-000000000000 -DisabledPlansMode Preserve
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "Medium")]
param(
    [string]$OutputFolder = (Join-Path $PWD "LicenseManagementOutput"),

    [string]$PilotUserCsv,

    [Guid]$PilotGroupId = [Guid]::Empty,

    [ValidateSet("Preserve", "None")]
    [string]$DisabledPlansMode = "Preserve",

    [switch]$IncludeDisabledAccounts,

    [switch]$IncludeInactiveSkus,

    [ValidateRange(0, 10)]
    [int]$MaxRetryCount = 4,

    [ValidateRange(1, 300)]
    [int]$RetryBaseDelaySeconds = 2,

    [ValidateRange(0, 60000)]
    [int]$RequestDelayMilliseconds = 250,

    [ValidateRange(0, 1000)]
    [int]$BatchSize = 20,

    [ValidateRange(0, 3600)]
    [int]$BatchPauseSeconds = 10
)

$ErrorActionPreference = "Stop"
$script:ScriptCmdlet = $PSCmdlet

# ---------------------------------------------------------------------------
# Initialization
# ---------------------------------------------------------------------------

if (-not (Test-Path -Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

$script:TimeStamp      = Get-Date -Format "yyyyMMdd_HHmmss"
$script:LogFile        = Join-Path $OutputFolder "LicenseManagement_$($script:TimeStamp).log"
$script:SubscribedSkus = @()
$script:SkuLookup      = @{}   # SkuId (string) -> SkuPartNumber
$script:SkuById        = @{}   # SkuId (string) -> SKU object
$script:CsvEncoding    = "utf8BOM"   # BOM so that Excel shows accented characters correctly

$script:UserProperties = @(
    "id",
    "displayName",
    "userPrincipalName",
    "mail",
    "accountEnabled",
    "userType",
    "usageLocation",
    "assignedLicenses",
    "licenseAssignmentStates"
)

function Write-LicenseLog {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $script:LogFile -Value $line -Encoding UTF8

    switch ($Level) {
        "INFO"    { Write-Host $Message -ForegroundColor Cyan }
        "WARNING" { Write-Host $Message -ForegroundColor Yellow }
        "ERROR"   { Write-Host $Message -ForegroundColor Red }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
    }
}

function Export-LicenseReport {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$InputObject,

        [Parameter(Mandatory)]
        [string]$Path
    )

    $InputObject |
        Export-Csv `
            -Path $Path `
            -NoTypeInformation `
            -Encoding $script:CsvEncoding `
            -Delimiter ";"
}

function Get-GraphErrorStatusCode {
    param(
        [Parameter(Mandatory)]
        [object]$ErrorRecord
    )

    $exception = $ErrorRecord.Exception

    foreach ($propertyName in @("ResponseStatusCode", "StatusCode")) {
        if ($exception.PSObject.Properties.Name -contains $propertyName -and $null -ne $exception.$propertyName) {
            return [int]$exception.$propertyName
        }
    }

    if ($exception.Response -and $exception.Response.StatusCode) {
        return [int]$exception.Response.StatusCode
    }

    return $null
}

function Get-GraphRetryAfterSeconds {
    param(
        [Parameter(Mandatory)]
        [object]$ErrorRecord
    )

    $exception = $ErrorRecord.Exception
    $headerCollections = @()

    foreach ($propertyName in @("ResponseHeaders", "Headers")) {
        if ($exception.PSObject.Properties.Name -contains $propertyName -and $null -ne $exception.$propertyName) {
            $headerCollections += $exception.$propertyName
        }
    }

    if ($exception.Response -and $exception.Response.Headers) {
        $headerCollections += $exception.Response.Headers
    }

    foreach ($headers in $headerCollections) {
        $retryAfter = $null

        if ($headers -is [System.Collections.IDictionary] -and $headers.Contains("Retry-After")) {
            $retryAfter = $headers["Retry-After"]
        }
        elseif ($headers.PSObject.Methods.Name -contains "TryGetValues") {
            $values = $null
            if ($headers.TryGetValues("Retry-After", [ref]$values)) {
                $retryAfter = @($values)[0]
            }
        }

        if ($retryAfter) {
            $seconds = 0
            if ([int]::TryParse([string]$retryAfter, [ref]$seconds)) {
                return [Math]::Max(1, $seconds)
            }
        }
    }

    return $null
}

function Test-GraphTransientError {
    param(
        [Parameter(Mandatory)]
        [object]$ErrorRecord
    )

    $statusCode = Get-GraphErrorStatusCode -ErrorRecord $ErrorRecord

    if ($statusCode -in @(408, 429, 500, 502, 503, 504)) {
        return $true
    }

    return ($ErrorRecord.Exception.Message -match "(?i)throttl|too many requests|timeout|temporar|service unavailable")
}

function Invoke-GraphRequest {
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory)]
        [string]$OperationName
    )

    for ($attempt = 1; $attempt -le ($MaxRetryCount + 1); $attempt++) {
        try {
            return & $ScriptBlock
        }
        catch {
            $isTransient = Test-GraphTransientError -ErrorRecord $_

            if (-not $isTransient -or $attempt -gt $MaxRetryCount) {
                throw
            }

            $retryAfter = Get-GraphRetryAfterSeconds -ErrorRecord $_
            $delay = if ($retryAfter) {
                $retryAfter
            }
            else {
                [Math]::Min(300, $RetryBaseDelaySeconds * [Math]::Pow(2, ($attempt - 1)))
            }

            $jitterMilliseconds = Get-Random -Minimum 0 -Maximum 1000

            Write-LicenseLog (
                "{0} failed with a transient Graph error. Retry {1}/{2} in {3:n1}s. Error: {4}" -f
                $OperationName,
                $attempt,
                $MaxRetryCount,
                ($delay + ($jitterMilliseconds / 1000)),
                $_.Exception.Message
            ) "WARNING"

            Start-Sleep -Seconds $delay

            if ($jitterMilliseconds -gt 0) {
                Start-Sleep -Milliseconds $jitterMilliseconds
            }
        }
    }
}

function Wait-LicenseRateLimit {
    param(
        [Parameter(Mandatory)]
        [int]$CompletedOperations,

        [Parameter(Mandatory)]
        [int]$TotalOperations
    )

    if ($CompletedOperations -ge $TotalOperations) {
        return
    }

    if ($RequestDelayMilliseconds -gt 0) {
        Start-Sleep -Milliseconds $RequestDelayMilliseconds
    }

    if ($BatchSize -gt 0 -and $BatchPauseSeconds -gt 0 -and ($CompletedOperations % $BatchSize) -eq 0) {
        Write-LicenseLog (
            "Rate limit pause: completed {0}/{1} write operations. Sleeping {2}s." -f
            $CompletedOperations,
            $TotalOperations,
            $BatchPauseSeconds
        ) "INFO"

        Start-Sleep -Seconds $BatchPauseSeconds
    }
}

function Initialize-GraphConnection {
    Write-LicenseLog "Checking Microsoft Graph PowerShell modules."

    $requiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Identity.DirectoryManagement",
        "Microsoft.Graph.Users",
        "Microsoft.Graph.Users.Actions"
    )

    if ($PilotGroupId -ne [Guid]::Empty) {
        $requiredModules += "Microsoft.Graph.Groups"
    }

    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-LicenseLog "Module $module not found. Installing it for the current user." "WARNING"

            Install-Module `
                -Name $module `
                -Scope CurrentUser `
                -Repository PSGallery `
                -Force
        }

        Import-Module -Name $module
    }

    $requiredScopes = @(
        "User.Read.All",
        "Organization.Read.All",
        "LicenseAssignment.ReadWrite.All"
    )

    if ($PilotGroupId -ne [Guid]::Empty) {
        $requiredScopes += "GroupMember.Read.All"
    }

    $context = Get-MgContext

    if (-not $context) {
        Write-LicenseLog "Connecting to Microsoft Graph."
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome
    }
    else {
        $missingScopes = @(
            $requiredScopes |
                Where-Object { $_ -notin $context.Scopes }
        )

        if ($missingScopes.Count -gt 0) {
            Write-LicenseLog ("The current session is missing these scopes: {0}. Re-authenticating." -f ($missingScopes -join ", ")) "WARNING"
            Disconnect-MgGraph | Out-Null
            Connect-MgGraph -Scopes $requiredScopes -NoWelcome
        }
    }

    $context = Get-MgContext

    if (-not $context) {
        throw "Unable to connect to Microsoft Graph."
    }

    Write-LicenseLog "Connected to tenant $($context.TenantId) as $($context.Account)." "SUCCESS"
    Update-SkuCache
}

# ---------------------------------------------------------------------------
# SKU helpers
# ---------------------------------------------------------------------------

function Update-SkuCache {
    Write-LicenseLog "Reading the SKUs subscribed by the tenant."

    $script:SubscribedSkus = @(
        Invoke-GraphRequest `
            -OperationName "Read subscribed SKUs" `
            -ScriptBlock { Get-MgSubscribedSku -All } |
            Sort-Object SkuPartNumber
    )

    $script:SkuLookup = @{}
    $script:SkuById   = @{}

    foreach ($sku in $script:SubscribedSkus) {
        $key = $sku.SkuId.ToString()
        $script:SkuLookup[$key] = $sku.SkuPartNumber
        $script:SkuById[$key]   = $sku
    }

    Write-LicenseLog "Found $($script:SubscribedSkus.Count) SKUs in the tenant." "SUCCESS"
}

function Get-SkuAvailability {
    param(
        [Parameter(Mandatory)]
        [object]$Sku
    )

    $enabled   = [int]$Sku.PrepaidUnits.Enabled
    $warning   = [int]$Sku.PrepaidUnits.Warning
    $suspended = [int]$Sku.PrepaidUnits.Suspended
    $consumed  = [int]$Sku.ConsumedUnits

    # Available = Enabled - Consumed (as in the Microsoft Graph documentation).
    # Warning units (subscription in grace period) are reported separately and
    # are NOT counted as available, to stay on the safe side.
    [PSCustomObject]@{
        EnabledUnits   = $enabled
        WarningUnits   = $warning
        SuspendedUnits = $suspended
        ConsumedUnits  = $consumed
        AvailableUnits = ($enabled - $consumed)
    }
}

function Get-UserServicePlanId {
    <#
        Returns the IDs of the user-level service plans of a SKU,
        optionally excluding a list of disabled plans.
    #>
    param(
        [Parameter(Mandatory)]
        [object]$Sku,

        [string[]]$DisabledPlanId = @()
    )

    @(
        $Sku.ServicePlans |
            Where-Object {
                $_.AppliesTo -eq "User" -and
                $_.ServicePlanId.ToString() -notin $DisabledPlanId
            } |
            ForEach-Object { $_.ServicePlanId.ToString() }
    )
}

function Get-ServicePlanName {
    param(
        [Parameter(Mandatory)]
        [string]$ServicePlanId
    )

    foreach ($sku in $script:SubscribedSkus) {
        $plan = $sku.ServicePlans |
            Where-Object { $_.ServicePlanId.ToString() -eq $ServicePlanId } |
            Select-Object -First 1

        if ($plan) {
            return $plan.ServicePlanName
        }
    }

    return $ServicePlanId
}

function Get-LicenseInventory {
    Update-SkuCache

    $inventory = foreach ($sku in $script:SubscribedSkus) {
        $availability = Get-SkuAvailability -Sku $sku

        [PSCustomObject]@{
            SkuPartNumber    = $sku.SkuPartNumber
            SkuId            = $sku.SkuId
            EnabledUnits     = $availability.EnabledUnits
            WarningUnits     = $availability.WarningUnits
            SuspendedUnits   = $availability.SuspendedUnits
            ConsumedUnits    = $availability.ConsumedUnits
            AvailableUnits   = $availability.AvailableUnits
            CapabilityStatus = $sku.CapabilityStatus
            AppliesTo        = $sku.AppliesTo
        }
    }

    return @($inventory)
}

function Show-LicenseInventory {
    Clear-Host
    Write-Host "LICENSES AVAILABLE IN THE TENANT" -ForegroundColor Green
    Write-Host ""

    $inventory = Get-LicenseInventory

    $inventory |
        Format-Table `
            SkuPartNumber,
            EnabledUnits,
            ConsumedUnits,
            AvailableUnits,
            WarningUnits,
            SuspendedUnits,
            CapabilityStatus `
            -AutoSize |
        Out-Host

    $exportAnswer = Read-Host "Export the inventory to CSV? [Y/N]"

    if ($exportAnswer -match "^[Yy]$") {
        $path = Join-Path $OutputFolder "TenantLicenses_$($script:TimeStamp).csv"
        Export-LicenseReport -InputObject $inventory -Path $path
        Write-LicenseLog "License inventory exported to $path." "SUCCESS"
    }
}

function Select-TenantSku {
    param(
        [Parameter(Mandatory)]
        [string]$Prompt,

        [Guid]$ExcludeSkuId = [Guid]::Empty
    )

    Update-SkuCache

    $availableSkus = @(
        $script:SubscribedSkus |
            Where-Object {
                ($ExcludeSkuId -eq [Guid]::Empty -or $_.SkuId -ne $ExcludeSkuId) -and
                ($IncludeInactiveSkus -or $_.CapabilityStatus -in @("Enabled", "Warning")) -and
                $_.AppliesTo -eq "User"
            } |
            Sort-Object SkuPartNumber
    )

    if ($availableSkus.Count -eq 0) {
        throw "No selectable SKU found in the tenant."
    }

    Write-Host ""
    Write-Host $Prompt -ForegroundColor Yellow
    Write-Host ""

    for ($index = 0; $index -lt $availableSkus.Count; $index++) {
        $sku          = $availableSkus[$index]
        $availability = Get-SkuAvailability -Sku $sku

        Write-Host (
            "[{0}] {1} | SKU: {2} | Assigned: {3} | Available: {4} | Status: {5}" -f
            ($index + 1),
            $sku.SkuPartNumber,
            $sku.SkuId,
            $availability.ConsumedUnits,
            $availability.AvailableUnits,
            $sku.CapabilityStatus
        )
    }

    $selectedIndex = 0

    do {
        $selection      = Read-Host "Enter the license number"
        $validSelection = [int]::TryParse($selection, [ref]$selectedIndex)

        if (
            -not $validSelection -or
            $selectedIndex -lt 1 -or
            $selectedIndex -gt $availableSkus.Count
        ) {
            Write-Host "Invalid selection." -ForegroundColor Red
            $validSelection = $false
        }
    }
    until ($validSelection)

    return $availableSkus[$selectedIndex - 1]
}

function ConvertTo-SkuPartNumberList {
    param(
        [object[]]$AssignedLicense
    )

    $names = foreach ($license in @($AssignedLicense)) {
        if (-not $license) { continue }

        $skuId = $license.SkuId.ToString()

        if ($script:SkuLookup.ContainsKey($skuId)) {
            $script:SkuLookup[$skuId]
        }
        else {
            "UNKNOWN_SKU:$skuId"
        }
    }

    return ($names -join "|")
}

# ---------------------------------------------------------------------------
# User helpers
# ---------------------------------------------------------------------------

function Get-PilotScope {
    <#
        Returns $null when no pilot is configured, otherwise an object with
        the set of allowed user IDs and UPNs.
    #>
    $hasCsv   = -not [string]::IsNullOrWhiteSpace($PilotUserCsv)
    $hasGroup = $PilotGroupId -ne [Guid]::Empty

    if (-not $hasCsv -and -not $hasGroup) {
        return $null
    }

    $upns = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $ids  = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

    if ($hasCsv) {
        if (-not (Test-Path -Path $PilotUserCsv)) {
            throw "Pilot CSV file not found: $PilotUserCsv"
        }

        $firstLine = Get-Content -Path $PilotUserCsv -TotalCount 1
        $delimiter = if ($firstLine -match ";") { ";" } else { "," }

        $rows = @(Import-Csv -Path $PilotUserCsv -Delimiter $delimiter)

        if ($rows.Count -gt 0 -and -not ($rows[0].PSObject.Properties.Name -contains "UserPrincipalName")) {
            throw "The pilot CSV must contain a 'UserPrincipalName' column."
        }

        foreach ($row in $rows) {
            if (-not [string]::IsNullOrWhiteSpace($row.UserPrincipalName)) {
                [void]$upns.Add($row.UserPrincipalName.Trim())
            }
        }

        Write-LicenseLog "Pilot CSV loaded: $($upns.Count) UPNs." "INFO"
    }

    if ($hasGroup) {
        $members = @(
            Invoke-GraphRequest `
                -OperationName "Read transitive members for pilot group $PilotGroupId" `
                -ScriptBlock { Get-MgGroupTransitiveMember -GroupId $PilotGroupId.ToString() -All }
        )

        foreach ($member in $members) {
            if ($member.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.user") {
                [void]$ids.Add($member.Id)
            }
        }

        Write-LicenseLog "Pilot group $PilotGroupId loaded: $($ids.Count) users." "INFO"
    }

    [PSCustomObject]@{
        Upns = $upns
        Ids  = $ids
    }
}

function Test-UserInPilotScope {
    param(
        [Parameter(Mandatory)]
        [object]$User,

        [object]$PilotScope
    )

    if ($null -eq $PilotScope) {
        return $true
    }

    return (
        $PilotScope.Ids.Contains([string]$User.Id) -or
        $PilotScope.Upns.Contains([string]$User.UserPrincipalName)
    )
}

function Get-SkuAssignedUser {
    param(
        [Parameter(Mandatory)]
        [Guid]$SkuId
    )

    Write-LicenseLog "Searching users with SKU $SkuId."

    $users = $null

    try {
        # Server-side filter: only users holding the SKU are returned.
        $users = @(
            Invoke-GraphRequest `
                -OperationName "Search users with SKU $SkuId" `
                -ScriptBlock {
                    Get-MgUser `
                        -All `
                        -Filter "assignedLicenses/any(x:x/skuId eq $SkuId)" `
                        -ConsistencyLevel eventual `
                        -CountVariable userCount `
                        -Property $script:UserProperties
                }
        )
    }
    catch {
        Write-LicenseLog "Server-side filter not available ($($_.Exception.Message)). Falling back to local filtering." "WARNING"

        $allUsers = Invoke-GraphRequest `
            -OperationName "Read all users for local SKU filtering" `
            -ScriptBlock { Get-MgUser -All -Property $script:UserProperties }

        $users = @(
            $allUsers |
                Where-Object {
                    $assignedSkuIds = @(
                        $_.AssignedLicenses |
                            ForEach-Object { $_.SkuId.ToString() }
                    )

                    $SkuId.ToString() -in $assignedSkuIds
                }
        )
    }

    Write-LicenseLog "Found $($users.Count) users with the selected SKU." "SUCCESS"

    return , $users
}

function Get-LicenseAssignmentInformation {
    param(
        [Parameter(Mandatory)]
        [object]$User,

        [Parameter(Mandatory)]
        [Guid]$SkuId
    )

    $states = @(
        $User.LicenseAssignmentStates |
            Where-Object {
                $_.SkuId.ToString() -eq $SkuId.ToString()
            }
    )

    $directAssignment = $false
    $groupIds         = @()
    $assignmentErrors = @()
    $assignmentStates = @()

    foreach ($state in $states) {
        if ([string]::IsNullOrWhiteSpace($state.AssignedByGroup)) {
            $directAssignment = $true
        }
        else {
            $groupIds += $state.AssignedByGroup
        }

        if (-not [string]::IsNullOrWhiteSpace($state.Error) -and $state.Error -ne "None") {
            $assignmentErrors += $state.Error
        }

        if (-not [string]::IsNullOrWhiteSpace($state.State)) {
            $assignmentStates += $state.State
        }
    }

    [PSCustomObject]@{
        IsDirect         = $directAssignment
        IsGroupBased     = ($groupIds.Count -gt 0)
        AssignedByGroups = (($groupIds | Sort-Object -Unique) -join "|")
        AssignmentStates = (($assignmentStates | Sort-Object -Unique) -join "|")
        AssignmentErrors = (($assignmentErrors | Sort-Object -Unique) -join "|")
    }
}

function Export-SkuAssignedUser {
    Clear-Host
    Write-Host "EXPORT USERS BY LICENSE" -ForegroundColor Green

    $selectedSku = Select-TenantSku `
        -Prompt "Select the license to filter on"

    $users = Get-SkuAssignedUser -SkuId $selectedSku.SkuId

    if ($users.Count -eq 0) {
        Write-LicenseLog "No users found with $($selectedSku.SkuPartNumber)." "WARNING"
        return
    }

    $report = foreach ($user in $users) {
        $assignment = Get-LicenseAssignmentInformation `
            -User $user `
            -SkuId $selectedSku.SkuId

        $selectedLicense = @($user.AssignedLicenses) |
            Where-Object { $_.SkuId.ToString() -eq $selectedSku.SkuId.ToString() } |
            Select-Object -First 1

        $disabledPlanNames = @(
            @($selectedLicense.DisabledPlans) |
                Where-Object { $_ } |
                ForEach-Object { Get-ServicePlanName -ServicePlanId $_.ToString() }
        ) -join "|"

        [PSCustomObject]@{
            DisplayName              = $user.DisplayName
            UserPrincipalName        = $user.UserPrincipalName
            Mail                     = $user.Mail
            AccountEnabled           = $user.AccountEnabled
            UserType                 = $user.UserType
            UsageLocation            = $user.UsageLocation
            SelectedSkuPartNumber    = $selectedSku.SkuPartNumber
            SelectedSkuId            = $selectedSku.SkuId
            SelectedLicenseIsDirect  = $assignment.IsDirect
            SelectedLicenseFromGroup = $assignment.IsGroupBased
            AssignedByGroupIds       = $assignment.AssignedByGroups
            LicenseAssignmentState   = $assignment.AssignmentStates
            LicenseAssignmentErrors  = $assignment.AssignmentErrors
            SelectedDisabledPlans    = $disabledPlanNames
            AllAssignedLicenses      = ConvertTo-SkuPartNumberList -AssignedLicense $user.AssignedLicenses
            AllAssignedSkuIds        = (@($user.AssignedLicenses) | ForEach-Object { $_.SkuId.ToString() }) -join "|"
        }
    }

    $path = Join-Path $OutputFolder (
        "UsersWith_{0}_{1}.csv" -f
        $selectedSku.SkuPartNumber,
        $script:TimeStamp
    )

    Export-LicenseReport -InputObject @($report | Sort-Object UserPrincipalName) -Path $path

    $report |
        Select-Object `
            DisplayName,
            UserPrincipalName,
            SelectedLicenseIsDirect,
            SelectedLicenseFromGroup,
            AllAssignedLicenses |
        Format-Table -AutoSize |
        Out-Host

    Write-LicenseLog "Export completed: $path." "SUCCESS"
}

# ---------------------------------------------------------------------------
# License replacement
# ---------------------------------------------------------------------------

function Invoke-LicenseReplacement {
    Clear-Host
    Write-Host "BULK LICENSE REPLACEMENT" -ForegroundColor Green
    Write-Host ""

    $pilotScope = Get-PilotScope

    if ($null -eq $pilotScope) {
        Write-LicenseLog "No pilot scope configured: ALL users holding the source license will be evaluated." "WARNING"
    }

    $sourceSku = Select-TenantSku `
        -Prompt "Select the SOURCE license to remove"

    $targetSku = Select-TenantSku `
        -Prompt "Select the TARGET license to assign" `
        -ExcludeSkuId $sourceSku.SkuId

    if ($sourceSku.SkuId -eq $targetSku.SkuId) {
        Write-LicenseLog "Source and target licenses are the same." "ERROR"
        return
    }

    # SKU-level service plan comparison.
    $sourcePlanIds = Get-UserServicePlanId -Sku $sourceSku
    $targetPlanIds = Get-UserServicePlanId -Sku $targetSku

    $skuLevelLostPlans = @(
        $sourcePlanIds |
            Where-Object { $_ -notin $targetPlanIds } |
            ForEach-Object { Get-ServicePlanName -ServicePlanId $_ } |
            Sort-Object
    )

    $allUsersWithSource = Get-SkuAssignedUser -SkuId $sourceSku.SkuId

    $users = @(
        $allUsersWithSource |
            Where-Object { Test-UserInPilotScope -User $_ -PilotScope $pilotScope }
    )

    if ($null -ne $pilotScope) {
        Write-LicenseLog "$($users.Count) of $($allUsersWithSource.Count) users holding the source license are in the pilot scope." "INFO"
    }

    if ($users.Count -eq 0) {
        Write-LicenseLog "No users in scope hold $($sourceSku.SkuPartNumber)." "WARNING"
        return
    }

    $preCheck = foreach ($user in $users) {
        $sourceAssignment = Get-LicenseAssignmentInformation `
            -User $user `
            -SkuId $sourceSku.SkuId

        $assignedLicenses = @($user.AssignedLicenses)
        $assignedSkuIds   = @($assignedLicenses | ForEach-Object { $_.SkuId.ToString() })

        $alreadyHasTarget = ($targetSku.SkuId.ToString() -in $assignedSkuIds)

        $sourceLicense = $assignedLicenses |
            Where-Object { $_.SkuId.ToString() -eq $sourceSku.SkuId.ToString() } |
            Select-Object -First 1

        $sourceDisabledPlanIds = @(
            @($sourceLicense.DisabledPlans) |
                Where-Object { $_ } |
                ForEach-Object { $_.ToString() }
        )

        # Disabled plans to carry over to the target (only those existing in the target).
        $targetDisabledPlanIds = @()

        if ($DisabledPlansMode -eq "Preserve" -and -not $alreadyHasTarget) {
            $targetDisabledPlanIds = @(
                $sourceDisabledPlanIds |
                    Where-Object { $_ -in $targetPlanIds }
            )
        }

        # Service plans the user would lose: enabled in source, not provided by
        # the target (as it will be assigned) nor by any other remaining SKU.
        $userSourceEnabledPlans = Get-UserServicePlanId `
            -Sku $sourceSku `
            -DisabledPlanId $sourceDisabledPlanIds

        $providedAfter = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

        if ($alreadyHasTarget) {
            $existingTarget = $assignedLicenses |
                Where-Object { $_.SkuId.ToString() -eq $targetSku.SkuId.ToString() } |
                Select-Object -First 1

            $existingTargetDisabled = @(@($existingTarget.DisabledPlans) | Where-Object { $_ } | ForEach-Object { $_.ToString() })

            Get-UserServicePlanId -Sku $targetSku -DisabledPlanId $existingTargetDisabled |
                ForEach-Object { [void]$providedAfter.Add($_) }
        }
        else {
            Get-UserServicePlanId -Sku $targetSku -DisabledPlanId $targetDisabledPlanIds |
                ForEach-Object { [void]$providedAfter.Add($_) }
        }

        foreach ($otherLicense in $assignedLicenses) {
            $otherId = $otherLicense.SkuId.ToString()

            if ($otherId -in @($sourceSku.SkuId.ToString(), $targetSku.SkuId.ToString())) { continue }
            if (-not $script:SkuById.ContainsKey($otherId)) { continue }

            $otherDisabled = @(@($otherLicense.DisabledPlans) | Where-Object { $_ } | ForEach-Object { $_.ToString() })

            Get-UserServicePlanId -Sku $script:SkuById[$otherId] -DisabledPlanId $otherDisabled |
                ForEach-Object { [void]$providedAfter.Add($_) }
        }

        $lostPlanNames = @(
            $userSourceEnabledPlans |
                Where-Object { -not $providedAfter.Contains($_) } |
                ForEach-Object { Get-ServicePlanName -ServicePlanId $_ } |
                Sort-Object
        )

        $status = "Ready"

        if ($sourceAssignment.IsGroupBased) {
            $status = "Skipped-SourceGroupBased"
        }
        elseif (-not $sourceAssignment.IsDirect) {
            $status = "Skipped-SourceNotDirect"
        }
        elseif ([string]::IsNullOrWhiteSpace($user.UsageLocation)) {
            $status = "Skipped-MissingUsageLocation"
        }
        elseif ($user.AccountEnabled -eq $false -and -not $IncludeDisabledAccounts) {
            $status = "Skipped-AccountDisabled"
        }
        elseif ($user.AccountEnabled -eq $false) {
            $status = "Ready-AccountDisabled"
        }

        [PSCustomObject]@{
            UserId                   = $user.Id
            DisplayName              = $user.DisplayName
            UserPrincipalName        = $user.UserPrincipalName
            AccountEnabled           = $user.AccountEnabled
            UsageLocation            = $user.UsageLocation
            SourceSkuPartNumber      = $sourceSku.SkuPartNumber
            SourceSkuId              = $sourceSku.SkuId
            TargetSkuPartNumber      = $targetSku.SkuPartNumber
            TargetSkuId              = $targetSku.SkuId
            SourceIsDirect           = $sourceAssignment.IsDirect
            SourceIsGroupBased       = $sourceAssignment.IsGroupBased
            SourceAssignedByGroups   = $sourceAssignment.AssignedByGroups
            AlreadyHasTarget         = $alreadyHasTarget
            SourceDisabledPlans      = (@($sourceDisabledPlanIds | ForEach-Object { Get-ServicePlanName -ServicePlanId $_ }) -join "|")
            TargetDisabledPlans      = (@($targetDisabledPlanIds | ForEach-Object { Get-ServicePlanName -ServicePlanId $_ }) -join "|")
            TargetDisabledPlanIds    = ($targetDisabledPlanIds -join "|")
            LostServicePlans         = ($lostPlanNames -join "|")
            LosesServices            = ($lostPlanNames.Count -gt 0)
            CurrentLicenses          = ConvertTo-SkuPartNumberList -AssignedLicense $assignedLicenses
            Status                   = $status
            # Original assignment kept for the JSON backup (not exported to CSV).
            OriginalAssignedLicenses = @(
                $assignedLicenses | ForEach-Object {
                    [PSCustomObject]@{
                        SkuId         = $_.SkuId.ToString()
                        SkuPartNumber = $script:SkuLookup[$_.SkuId.ToString()]
                        DisabledPlans = @(@($_.DisabledPlans) | Where-Object { $_ } | ForEach-Object { $_.ToString() })
                    }
                }
            )
        }
    }

    $preCheck = @($preCheck)

    $preCheckPath = Join-Path $OutputFolder (
        "LicenseReplacement_PreCheck_{0}_to_{1}_{2}.csv" -f
        $sourceSku.SkuPartNumber,
        $targetSku.SkuPartNumber,
        $script:TimeStamp
    )

    Export-LicenseReport `
        -InputObject @($preCheck | Select-Object -Property * -ExcludeProperty OriginalAssignedLicenses) `
        -Path $preCheckPath

    $eligibleUsers = @(
        $preCheck |
            Where-Object { $_.Status -in @("Ready", "Ready-AccountDisabled") }
    )

    $groupBasedUsers      = @($preCheck | Where-Object { $_.SourceIsGroupBased -eq $true })
    $missingLocationUsers = @($preCheck | Where-Object { $_.Status -eq "Skipped-MissingUsageLocation" })
    $disabledSkippedUsers = @($preCheck | Where-Object { $_.Status -eq "Skipped-AccountDisabled" })
    $usersLosingServices  = @($eligibleUsers | Where-Object { $_.LosesServices -eq $true })

    $targetAvailability = Get-SkuAvailability -Sku $targetSku

    $newTargetAssignmentsRequired = @(
        $eligibleUsers |
            Where-Object { $_.AlreadyHasTarget -eq $false }
    ).Count

    Write-Host ""
    Write-Host "Pre-check summary" -ForegroundColor Yellow
    Write-Host "Source license             : $($sourceSku.SkuPartNumber)"
    Write-Host "Target license             : $($targetSku.SkuPartNumber)"
    Write-Host "Disabled plans mode        : $DisabledPlansMode"
    Write-Host "Pilot scope                : $(if ($null -eq $pilotScope) { 'NONE (all users)' } else { 'ACTIVE' })"
    Write-Host "Users with source in scope : $($users.Count)"
    Write-Host "Users to be changed        : $($eligibleUsers.Count)"
    Write-Host "Target licenses needed     : $newTargetAssignmentsRequired"
    Write-Host "Target licenses available  : $($targetAvailability.AvailableUnits) (warning units not counted: $($targetAvailability.WarningUnits))"
    Write-Host "Skipped - source by group  : $($groupBasedUsers.Count)"
    Write-Host "Skipped - no UsageLocation : $($missingLocationUsers.Count)"
    Write-Host "Skipped - disabled account : $($disabledSkippedUsers.Count)"
    Write-Host "Users losing services      : $($usersLosingServices.Count)"
    Write-Host "Pre-check report           : $preCheckPath"
    Write-Host ""

    if ($skuLevelLostPlans.Count -gt 0) {
        Write-LicenseLog (
            "Service plans included in {0} but NOT in {1}: {2}" -f
            $sourceSku.SkuPartNumber,
            $targetSku.SkuPartNumber,
            ($skuLevelLostPlans -join ", ")
        ) "WARNING"
    }

    if ($eligibleUsers.Count -eq 0) {
        Write-LicenseLog "No users can be changed directly." "WARNING"
        return
    }

    if ($newTargetAssignmentsRequired -gt $targetAvailability.AvailableUnits) {
        Write-LicenseLog (
            "Not enough target licenses. Required: {0}; available: {1}." -f
            $newTargetAssignmentsRequired,
            $targetAvailability.AvailableUnits
        ) "ERROR"

        return
    }

    if ($groupBasedUsers.Count -gt 0) {
        Write-LicenseLog ("{0} users skipped because the source license is assigned by a group." -f $groupBasedUsers.Count) "WARNING"
    }

    if ($missingLocationUsers.Count -gt 0) {
        Write-LicenseLog ("{0} users skipped because UsageLocation is empty." -f $missingLocationUsers.Count) "WARNING"
    }

    if ($disabledSkippedUsers.Count -gt 0) {
        Write-LicenseLog ("{0} disabled accounts skipped (use -IncludeDisabledAccounts to process them)." -f $disabledSkippedUsers.Count) "WARNING"
    }

    if ($usersLosingServices.Count -gt 0) {
        Write-LicenseLog (
            "{0} users would LOSE at least one service plan. Check the 'LostServicePlans' column of the pre-check report before continuing." -f
            $usersLosingServices.Count
        ) "WARNING"
    }

    if ($WhatIfPreference) {
        Write-LicenseLog "WhatIf is active: pre-check report created; no backup or license changes will be made." "WARNING"
        return
    }

    # JSON backup of the current assignment, written BEFORE any change.
    $backupPath = Join-Path $OutputFolder (
        "LicenseReplacement_Backup_{0}_to_{1}_{2}.json" -f
        $sourceSku.SkuPartNumber,
        $targetSku.SkuPartNumber,
        $script:TimeStamp
    )

    $backup = [PSCustomObject]@{
        CreatedAt           = (Get-Date).ToString("o")
        TenantId            = (Get-MgContext).TenantId
        SourceSkuId         = $sourceSku.SkuId.ToString()
        SourceSkuPartNumber = $sourceSku.SkuPartNumber
        TargetSkuId         = $targetSku.SkuId.ToString()
        TargetSkuPartNumber = $targetSku.SkuPartNumber
        DisabledPlansMode   = $DisabledPlansMode
        Users               = @(
            $eligibleUsers | ForEach-Object {
                [PSCustomObject]@{
                    UserId                   = $_.UserId
                    UserPrincipalName        = $_.UserPrincipalName
                    TargetAddedByScript      = (-not $_.AlreadyHasTarget)
                    OriginalAssignedLicenses = $_.OriginalAssignedLicenses
                }
            }
        )
    }

    $backup | ConvertTo-Json -Depth 6 | Set-Content -Path $backupPath -Encoding UTF8
    Write-LicenseLog "Backup of current assignments saved to $backupPath." "SUCCESS"

    Write-Host ""
    Write-Host "SIMULATION MODE" -ForegroundColor Cyan
    Write-Host "No change will be made until you type EXECUTE."
    Write-Host ""

    $confirmation = Read-Host (
        "Type EXECUTE to replace {0} with {1} on {2} users" -f
        $sourceSku.SkuPartNumber,
        $targetSku.SkuPartNumber,
        $eligibleUsers.Count
    )

    if ($confirmation -cne "EXECUTE") {
        Write-LicenseLog "Operation cancelled. Only the pre-check and the backup were produced." "WARNING"
        return
    }

    Write-LicenseLog (
        "Replacement started by {0}: {1} -> {2} on {3} users." -f
        (Get-MgContext).Account,
        $sourceSku.SkuPartNumber,
        $targetSku.SkuPartNumber,
        $eligibleUsers.Count
    ) "INFO"

    $counter = 0

    $result = foreach ($item in $eligibleUsers) {
        $counter++
        $startedAt = Get-Date

        Write-Progress `
            -Activity "Replacing $($sourceSku.SkuPartNumber) with $($targetSku.SkuPartNumber)" `
            -Status "$counter / $($eligibleUsers.Count) - $($item.UserPrincipalName)" `
            -PercentComplete (($counter / $eligibleUsers.Count) * 100)

        try {
            Write-LicenseLog (
                "Processing {0}: {1} -> {2}." -f
                $item.UserPrincipalName,
                $sourceSku.SkuPartNumber,
                $targetSku.SkuPartNumber
            )

            if ($item.AlreadyHasTarget) {
                # The user already holds the target: remove only the source.
                if (-not $script:ScriptCmdlet.ShouldProcess($item.UserPrincipalName, "Remove source license $($sourceSku.SkuPartNumber)")) {
                    [PSCustomObject]@{
                        DisplayName         = $item.DisplayName
                        UserPrincipalName   = $item.UserPrincipalName
                        SourceSku           = $sourceSku.SkuPartNumber
                        TargetSku           = $targetSku.SkuPartNumber
                        Operation           = "SkippedByShouldProcess"
                        Status              = "Skipped"
                        SourceStillAssigned = $null
                        TargetAssigned      = $null
                        MissingOtherSkus    = $null
                        TargetDisabledPlans = $item.TargetDisabledPlans
                        LostServicePlans    = $item.LostServicePlans
                        FinalLicenses       = $null
                        Error               = $null
                        StartedAt           = $startedAt
                        CompletedAt         = Get-Date
                    }

                    continue
                }

                Invoke-GraphRequest `
                    -OperationName "Remove source license from $($item.UserPrincipalName)" `
                    -ScriptBlock {
                        Set-MgUserLicense `
                            -UserId $item.UserId `
                            -AddLicenses @() `
                            -RemoveLicenses @($sourceSku.SkuId) `
                            -Confirm:$false |
                            Out-Null
                    }

                $operation = "RemovedSource-TargetAlreadyAssigned"
            }
            else {
                # Single call: add the target and remove only the source.
                $disabledPlans = @(
                    $item.TargetDisabledPlanIds -split "\|" |
                        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                )

                $licenseToAdd = @(
                    @{
                        SkuId         = $targetSku.SkuId
                        DisabledPlans = $disabledPlans
                    }
                )

                if (-not $script:ScriptCmdlet.ShouldProcess($item.UserPrincipalName, "Add target license $($targetSku.SkuPartNumber) and remove source license $($sourceSku.SkuPartNumber)")) {
                    [PSCustomObject]@{
                        DisplayName         = $item.DisplayName
                        UserPrincipalName   = $item.UserPrincipalName
                        SourceSku           = $sourceSku.SkuPartNumber
                        TargetSku           = $targetSku.SkuPartNumber
                        Operation           = "SkippedByShouldProcess"
                        Status              = "Skipped"
                        SourceStillAssigned = $null
                        TargetAssigned      = $null
                        MissingOtherSkus    = $null
                        TargetDisabledPlans = $item.TargetDisabledPlans
                        LostServicePlans    = $item.LostServicePlans
                        FinalLicenses       = $null
                        Error               = $null
                        StartedAt           = $startedAt
                        CompletedAt         = Get-Date
                    }

                    continue
                }

                Invoke-GraphRequest `
                    -OperationName "Replace license for $($item.UserPrincipalName)" `
                    -ScriptBlock {
                        Set-MgUserLicense `
                            -UserId $item.UserId `
                            -AddLicenses $licenseToAdd `
                            -RemoveLicenses @($sourceSku.SkuId) `
                            -Confirm:$false |
                            Out-Null
                    }

                $operation = "AddedTarget-And-RemovedSource"
            }

            # Post-operation verification.
            $updatedUser = Invoke-GraphRequest `
                -OperationName "Verify license replacement for $($item.UserPrincipalName)" `
                -ScriptBlock {
                    Get-MgUser `
                        -UserId $item.UserId `
                        -Property @("id", "displayName", "userPrincipalName", "assignedLicenses")
                }

            $updatedSkuIds = @(
                $updatedUser.AssignedLicenses |
                    ForEach-Object { $_.SkuId.ToString() }
            )

            $sourceStillAssigned = ($sourceSku.SkuId.ToString() -in $updatedSkuIds)
            $targetAssigned      = ($targetSku.SkuId.ToString() -in $updatedSkuIds)

            # Every other SKU must still be there.
            $originalOtherSkuIds = @(
                $item.OriginalAssignedLicenses |
                    Where-Object { $_.SkuId -ne $sourceSku.SkuId.ToString() } |
                    ForEach-Object { $_.SkuId }
            )

            $missingOtherSkus = @(
                $originalOtherSkuIds |
                    Where-Object { $_ -notin $updatedSkuIds }
            )

            if (-not $sourceStillAssigned -and $targetAssigned -and $missingOtherSkus.Count -eq 0) {
                $finalStatus = "Success"
                Write-LicenseLog "Completed for $($item.UserPrincipalName)." "SUCCESS"
            }
            else {
                $finalStatus = "VerificationFailed"
                Write-LicenseLog "Verification failed for $($item.UserPrincipalName)." "ERROR"
            }

            [PSCustomObject]@{
                DisplayName         = $item.DisplayName
                UserPrincipalName   = $item.UserPrincipalName
                SourceSku           = $sourceSku.SkuPartNumber
                TargetSku           = $targetSku.SkuPartNumber
                Operation           = $operation
                Status              = $finalStatus
                SourceStillAssigned = $sourceStillAssigned
                TargetAssigned      = $targetAssigned
                MissingOtherSkus    = (@($missingOtherSkus | ForEach-Object { $script:SkuLookup[$_] }) -join "|")
                TargetDisabledPlans = $item.TargetDisabledPlans
                LostServicePlans    = $item.LostServicePlans
                FinalLicenses       = ConvertTo-SkuPartNumberList -AssignedLicense $updatedUser.AssignedLicenses
                Error               = $null
                StartedAt           = $startedAt
                CompletedAt         = Get-Date
            }

            Wait-LicenseRateLimit -CompletedOperations $counter -TotalOperations $eligibleUsers.Count
        }
        catch {
            Write-LicenseLog (
                "Error on {0}: {1}" -f
                $item.UserPrincipalName,
                $_.Exception.Message
            ) "ERROR"

            [PSCustomObject]@{
                DisplayName         = $item.DisplayName
                UserPrincipalName   = $item.UserPrincipalName
                SourceSku           = $sourceSku.SkuPartNumber
                TargetSku           = $targetSku.SkuPartNumber
                Operation           = "Failed"
                Status              = "Error"
                SourceStillAssigned = $null
                TargetAssigned      = $null
                MissingOtherSkus    = $null
                TargetDisabledPlans = $item.TargetDisabledPlans
                LostServicePlans    = $item.LostServicePlans
                FinalLicenses       = $null
                Error               = $_.Exception.Message
                StartedAt           = $startedAt
                CompletedAt         = Get-Date
            }
        }
    }

    Write-Progress -Activity "License replacement" -Completed

    $result = @($result)

    $resultPath = Join-Path $OutputFolder (
        "LicenseReplacement_Result_{0}_to_{1}_{2}.csv" -f
        $sourceSku.SkuPartNumber,
        $targetSku.SkuPartNumber,
        $script:TimeStamp
    )

    Export-LicenseReport -InputObject $result -Path $resultPath

    $successCount = @($result | Where-Object { $_.Status -eq "Success" }).Count
    $failureCount = @($result | Where-Object { $_.Status -ne "Success" }).Count

    Write-Host ""
    Write-Host "Final result" -ForegroundColor Yellow
    Write-Host "Succeeded      : $successCount"
    Write-Host "Failed         : $failureCount"
    Write-Host "Result report  : $resultPath"
    Write-Host "Backup (JSON)  : $backupPath"
    Write-Host "Log file       : $($script:LogFile)"

    if ($failureCount -eq 0) {
        Write-LicenseLog "Replacement completed without errors." "SUCCESS"
    }
    else {
        Write-LicenseLog "Replacement completed with $failureCount errors." "WARNING"
    }
}

# ---------------------------------------------------------------------------
# Rollback
# ---------------------------------------------------------------------------

function Restore-LicenseBackup {
    Clear-Host
    Write-Host "RESTORE LICENSES FROM BACKUP" -ForegroundColor Green
    Write-Host ""

    $backupFiles = @(
        Get-ChildItem -Path $OutputFolder -Filter "LicenseReplacement_Backup_*.json" |
            Sort-Object LastWriteTime -Descending
    )

    if ($backupFiles.Count -eq 0) {
        Write-LicenseLog "No backup file found in $OutputFolder." "WARNING"
        return
    }

    for ($index = 0; $index -lt $backupFiles.Count; $index++) {
        Write-Host ("[{0}] {1}" -f ($index + 1), $backupFiles[$index].Name)
    }

    $selectedIndex = 0

    do {
        $selection      = Read-Host "Enter the backup number"
        $validSelection = [int]::TryParse($selection, [ref]$selectedIndex)

        if (-not $validSelection -or $selectedIndex -lt 1 -or $selectedIndex -gt $backupFiles.Count) {
            Write-Host "Invalid selection." -ForegroundColor Red
            $validSelection = $false
        }
    }
    until ($validSelection)

    $backupFile = $backupFiles[$selectedIndex - 1]
    $backup     = Get-Content -Path $backupFile.FullName -Raw | ConvertFrom-Json

    if ($backup.TenantId -ne (Get-MgContext).TenantId) {
        Write-LicenseLog "The backup belongs to tenant $($backup.TenantId), not to the current tenant." "ERROR"
        return
    }

    $sourceSkuId = $backup.SourceSkuId
    $targetSkuId = $backup.TargetSkuId
    $users       = @($backup.Users)

    Write-Host ""
    Write-Host "Backup     : $($backupFile.Name)"
    Write-Host "Restore    : $($backup.TargetSkuPartNumber) -> $($backup.SourceSkuPartNumber)"
    Write-Host "Users      : $($users.Count)"
    Write-Host ""
    Write-Host "The source license will be re-assigned with its original disabled plans."
    Write-Host "The target license will be removed ONLY where it was added by the script."
    Write-Host ""

    if ($WhatIfPreference) {
        Write-LicenseLog "WhatIf is active: restore preview completed; no license changes will be made." "WARNING"
        return
    }

    $confirmation = Read-Host "Type RESTORE to proceed"

    if ($confirmation -cne "RESTORE") {
        Write-LicenseLog "Restore cancelled." "WARNING"
        return
    }

    $counter = 0

    $result = foreach ($user in $users) {
        $counter++

        try {
            $originalSource = @($user.OriginalAssignedLicenses) |
                Where-Object { $_.SkuId -eq $sourceSkuId } |
                Select-Object -First 1

            $licenseToAdd = @(
                @{
                    SkuId         = $sourceSkuId
                    DisabledPlans = @(@($originalSource.DisabledPlans) | Where-Object { $_ })
                }
            )

            $licensesToRemove = @()

            if ($user.TargetAddedByScript) {
                $licensesToRemove = @($targetSkuId)
            }

            if (-not $script:ScriptCmdlet.ShouldProcess($user.UserPrincipalName, "Restore source license $($backup.SourceSkuPartNumber) from backup")) {
                [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    Status            = "Skipped"
                    Error             = $null
                }

                continue
            }

            Invoke-GraphRequest `
                -OperationName "Restore licenses for $($user.UserPrincipalName)" `
                -ScriptBlock {
                    Set-MgUserLicense `
                        -UserId $user.UserId `
                        -AddLicenses $licenseToAdd `
                        -RemoveLicenses $licensesToRemove `
                        -Confirm:$false |
                        Out-Null
                }

            Write-LicenseLog "Restored $($user.UserPrincipalName)." "SUCCESS"

            [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                Status            = "Restored"
                Error             = $null
            }
        }
        catch {
            Write-LicenseLog "Restore error on $($user.UserPrincipalName): $($_.Exception.Message)" "ERROR"

            [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                Status            = "Error"
                Error             = $_.Exception.Message
            }
        }

        Wait-LicenseRateLimit -CompletedOperations $counter -TotalOperations $users.Count
    }

    $resultPath = Join-Path $OutputFolder ("LicenseRestore_Result_{0}.csv" -f $script:TimeStamp)
    Export-LicenseReport -InputObject @($result) -Path $resultPath

    Write-LicenseLog "Restore completed. Report: $resultPath." "SUCCESS"
}

# ---------------------------------------------------------------------------
# Menu
# ---------------------------------------------------------------------------

function Show-MainMenu {
    do {
        Clear-Host

        Write-Host "==============================================" -ForegroundColor DarkCyan
        Write-Host "  MICROSOFT 365 LICENSE MANAGEMENT" -ForegroundColor Green
        Write-Host "==============================================" -ForegroundColor DarkCyan
        Write-Host ""
        Write-Host "1) Show the licenses available in the tenant"
        Write-Host "2) Export the users holding a license"
        Write-Host "3) Replace a license with another one"
        Write-Host "4) Restore licenses from a backup"
        Write-Host "5) Exit"
        Write-Host ""

        $choice = Read-Host "Select an option"

        try {
            switch ($choice) {
                "1" {
                    Show-LicenseInventory
                    Read-Host "Press ENTER to return to the menu"
                }

                "2" {
                    Export-SkuAssignedUser
                    Read-Host "Press ENTER to return to the menu"
                }

                "3" {
                    Invoke-LicenseReplacement
                    Read-Host "Press ENTER to return to the menu"
                }

                "4" {
                    Restore-LicenseBackup
                    Read-Host "Press ENTER to return to the menu"
                }

                "5" {
                    Write-LicenseLog "Closing the script."
                }

                default {
                    Write-Host "Invalid option." -ForegroundColor Red
                    Start-Sleep -Seconds 1
                }
            }
        }
        catch {
            Write-LicenseLog $_.Exception.Message "ERROR"
            Read-Host "Press ENTER to return to the menu"
        }
    }
    until ($choice -eq "5")
}

# ---------------------------------------------------------------------------
# Start
# ---------------------------------------------------------------------------

try {
    Initialize-GraphConnection
    Show-MainMenu
}
catch {
    Write-LicenseLog "Unhandled error: $($_.Exception.Message)" "ERROR"
    throw
}
finally {
    if (Get-MgContext) {
        Disconnect-MgGraph | Out-Null
    }
}
