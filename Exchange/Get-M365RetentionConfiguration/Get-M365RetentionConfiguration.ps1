<#
.SYNOPSIS
    Documents Microsoft 365 retention and hold configuration.

.DESCRIPTION
    This script collects and exports:

    - PowerShell and ExchangeOnlineManagement runtime information
    - Exchange Online organization configuration
    - Organization-wide holds
    - Microsoft Purview retention policies and rules
    - Microsoft Purview app retention policies
    - eDiscovery case hold policies
    - Exchange MRM retention policies and tags
    - Mailbox-level retention and hold settings
    - Decoded mailbox and organization InPlaceHolds
    - Effective cmdlet availability and RBAC read-access results

    The script does not make configuration changes.

.REQUIREMENTS
    - PowerShell 7.0 or later
    - PowerShell 7.4 or later recommended
    - ExchangeOnlineManagement 3.10.1 or later
    - Exchange Online read permissions
    - Microsoft Purview read permissions for compliance reporting

.NOTES
    Microsoft Learn reference:
    https://learn.microsoft.com/en-us/purview/ediscovery-identify-a-hold-on-an-exchange-online-mailbox

    Current documentation:
    https://learn.microsoft.com/en-us/purview/edisc-hold-types-mailboxes

    The RBAC preflight validates effective read access by running
    non-destructive read operations. It does not attempt to reconstruct
    the complete RBAC assignment chain.

    Organization-wide retention policies might apply to a mailbox even
    when the mailbox InPlaceHolds property is empty.

    Policy and rule reports are always arrays so summary counts also
    work for empty and single-item results under StrictMode.

    eDiscovery hold policies are queried per visible Standard and Premium
    case. Case visibility depends on the connected account's permissions;
    an empty result does not prove that the tenant has no case holds.
    Optional preflight failures produce a partial-access assessment.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = ".\M365-Retention-Assessment",

    [Parameter()]
    [string]$UserPrincipalName,

    [Parameter()]
    [switch]$IncludeSoftDeletedMailboxes,

    [Parameter()]
    [switch]$SkipComplianceConnection,

    [Parameter()]
    [switch]$ContinueOnPreflightFailure
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region General helper functions

function Write-RetentionLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter()]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Information `
        -MessageData "[$timestamp][$Level] $Message" `
        -InformationAction Continue
}

function Convert-ToSafeFileName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    $safeName = $Name

    foreach ($character in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safeName = $safeName.Replace($character, "_")
    }

    return $safeName
}

function Convert-ToFlatString {
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowNull()]
        $Value,

        [Parameter()]
        [string]$Separator = "; "
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [string]) {
        return $Value
    }

    if ($Value -is [System.Collections.IDictionary]) {
        return ($Value | ConvertTo-Json -Compress -Depth 10)
    }

    if ($Value -is [System.Collections.IEnumerable]) {
        return (
            @(
                $Value | ForEach-Object {
                    [string]$_
                }
            ) -join $Separator
        )
    }

    return [string]$Value
}

function Get-SafePropertyValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        $InputObject,

        [Parameter(Mandatory)]
        [string]$PropertyName,

        [Parameter()]
        [AllowNull()]
        $DefaultValue = $null
    )

    if ($null -eq $InputObject) {
        return $DefaultValue
    }

    $property = $InputObject.PSObject.Properties[$PropertyName]

    if ($null -eq $property) {
        return $DefaultValue
    }

    return $property.Value
}

function Export-DataSet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter()]
        [AllowNull()]
        $Data,

        [Parameter(Mandatory)]
        [string]$Destination
    )

    $safeName = Convert-ToSafeFileName -Name $Name
    $csvPath = Join-Path -Path $Destination -ChildPath "$safeName.csv"
    $jsonPath = Join-Path -Path $Destination -ChildPath "$safeName.json"
    $dataArray = @($Data)

    if ($dataArray.Count -gt 0) {
        $dataArray |
            Export-Csv `
                -Path $csvPath `
                -NoTypeInformation `
                -Encoding UTF8

        $dataArray |
            ConvertTo-Json -Depth 20 |
            Set-Content `
                -Path $jsonPath `
                -Encoding UTF8
    }
    else {
        Set-Content `
            -Path $csvPath `
            -Value "" `
            -Encoding UTF8

        "[]" |
            Set-Content `
                -Path $jsonPath `
                -Encoding UTF8
    }
}

function Test-PreflightPassed {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$PreflightResults,

        [Parameter(Mandatory)]
        [string]$TestName
    )

    $result = $PreflightResults |
        Where-Object {
            $_.Test -eq $TestName
        } |
        Select-Object -First 1

    return (
        $null -ne $result -and
        $result.Status -eq "Passed"
    )
}

#endregion

#region Runtime validation

function Test-RuntimeRequirement {
    [CmdletBinding()]
    param()

    $results = [System.Collections.Generic.List[object]]::new()

    $minimumPowerShellVersion = [version]"7.0.0"
    $recommendedPowerShellVersion = [version]"7.4.0"
    $minimumModuleVersion = [version]"3.10.1"

    $powerShellVersion = $PSVersionTable.PSVersion
    $powerShellPassed = (
        $powerShellVersion -ge $minimumPowerShellVersion
    )

    $results.Add(
        [PSCustomObject]@{
            Category = "Runtime"
            Service  = "PowerShell"
            Test     = "PowerShell version"
            Required = $minimumPowerShellVersion.ToString()
            Detected = $powerShellVersion.ToString()
            Status   = if ($powerShellPassed) {
                "Passed"
            }
            else {
                "Failed"
            }
            Blocking = $true
            Details  = if (-not $powerShellPassed) {
                "PowerShell 7.0 or later is required."
            }
            elseif (
                $powerShellVersion -lt $recommendedPowerShellVersion
            ) {
                "The version is supported, but PowerShell 7.4 or later is recommended."
            }
            else {
                "A supported and recommended PowerShell version is installed."
            }
        }
    )

    $exoModule = Get-Module `
        -ListAvailable `
        -Name ExchangeOnlineManagement |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($null -eq $exoModule) {
        $results.Add(
            [PSCustomObject]@{
                Category = "Runtime"
                Service  = "Exchange Online"
                Test     = "ExchangeOnlineManagement module"
                Required = $minimumModuleVersion.ToString()
                Detected = $null
                Status   = "Failed"
                Blocking = $true
                Details  = "The ExchangeOnlineManagement module is not installed."
            }
        )

        return $results
    }

    $modulePassed = (
        $exoModule.Version -ge $minimumModuleVersion
    )

    $results.Add(
        [PSCustomObject]@{
            Category = "Runtime"
            Service  = "Exchange Online"
            Test     = "ExchangeOnlineManagement module"
            Required = $minimumModuleVersion.ToString()
            Detected = $exoModule.Version.ToString()
            Status   = if ($modulePassed) {
                "Passed"
            }
            else {
                "Failed"
            }
            Blocking = $true
            Details  = if ($modulePassed) {
                "A supported ExchangeOnlineManagement version is installed."
            }
            else {
                "Upgrade ExchangeOnlineManagement before running the assessment."
            }
        }
    )

    if ($powerShellPassed -and $modulePassed) {
        Import-Module ExchangeOnlineManagement `
            -MinimumVersion $minimumModuleVersion `
            -ErrorAction Stop
    }

    return $results
}

#endregion

#region Cmdlet and RBAC validation

function Get-VisibleCaseHoldPolicy {
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$RecordFailures
    )

    $seenPolicyGuids = @{}
    $failedCaseTypeQueries = @()
    $successfulCaseTypeQueries = 0

    foreach ($caseType in @("eDiscovery", "AdvancedEdiscovery")) {
        try {
            $cases = @(
                Get-ComplianceCase -CaseType $caseType -ErrorAction Stop
            )
            $successfulCaseTypeQueries++

            foreach ($case in $cases) {
                $caseIdentity = Get-SafePropertyValue `
                    -InputObject $case `
                    -PropertyName "Identity"

                if ([string]::IsNullOrWhiteSpace([string]$caseIdentity)) {
                    $caseIdentity = Get-SafePropertyValue `
                        -InputObject $case `
                        -PropertyName "Guid"
                }

                if ([string]::IsNullOrWhiteSpace([string]$caseIdentity)) {
                    $failedCaseTypeQueries += "A visible $caseType case has no Identity or Guid."
                    continue
                }

                try {
                    $policies = @(
                        Get-CaseHoldPolicy -Case $caseIdentity -ErrorAction Stop
                    )

                    foreach ($policy in $policies) {
                        $policyGuid = Get-SafePropertyValue `
                            -InputObject $policy `
                            -PropertyName "Guid"

                        $policyKey = if ([string]::IsNullOrWhiteSpace([string]$policyGuid)) {
                            Convert-ToFlatString -Value @(
                                Get-SafePropertyValue -InputObject $policy -PropertyName "Name"
                                Get-SafePropertyValue -InputObject $policy -PropertyName "Identity"
                            )
                        }
                        else {
                            ([string]$policyGuid).ToLowerInvariant()
                        }

                        if (
                            [string]::IsNullOrWhiteSpace([string]$policyKey) -or
                            -not $seenPolicyGuids.ContainsKey([string]$policyKey)
                        ) {
                            if (-not [string]::IsNullOrWhiteSpace([string]$policyKey)) {
                                $seenPolicyGuids[[string]$policyKey] = $true
                            }
                            $policy
                        }
                    }
                }
                catch {
                    $message = "Unable to query hold policies for $caseType case '$caseIdentity': $($_.Exception.Message)"
                    $failedCaseTypeQueries += $message
                    if ($RecordFailures) {
                        $script:collectionFailures.Add(
                            [PSCustomObject]@{
                                Section = "eDiscovery case hold policies"
                                Details = $message
                            }
                        )
                    }
                }
            }
        }
        catch {
            $message = "Unable to query $caseType compliance cases: $($_.Exception.Message)"
            $failedCaseTypeQueries += $message
            if ($RecordFailures) {
                $script:collectionFailures.Add(
                    [PSCustomObject]@{
                        Section = "eDiscovery case hold policies"
                        Details = $message
                    }
                )
            }
        }
    }

    if (
        $successfulCaseTypeQueries -eq 0 -and
        $failedCaseTypeQueries.Count -gt 0
    ) {
        throw ($failedCaseTypeQueries -join " | ")
    }
}

function Test-CommandAvailability {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CommandName,

        [Parameter(Mandatory)]
        [ValidateSet("Exchange Online", "Microsoft Purview")]
        [string]$Service,

        [Parameter()]
        [bool]$Required = $true
    )

    $command = Get-Command `
        -Name $CommandName `
        -ErrorAction SilentlyContinue

    if ($null -eq $command) {
        return [PSCustomObject]@{
            Category   = "Cmdlet"
            Service    = $Service
            Test       = $CommandName
            Status     = "Failed"
            Blocking   = $Required
            RbacStatus = "Unknown"
            Details    = "The cmdlet is unavailable. Verify the connection, module version, and RBAC permissions."
        }
    }

    return [PSCustomObject]@{
        Category   = "Cmdlet"
        Service    = $Service
        Test       = $CommandName
        Status     = "Passed"
        Blocking   = $false
        RbacStatus = "NotTested"
        Details    = "The cmdlet is available in the current session."
    }
}

function Test-ReadAccess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TestName,

        [Parameter(Mandatory)]
        [ValidateSet("Exchange Online", "Microsoft Purview")]
        [string]$Service,

        [Parameter(Mandatory)]
        [scriptblock]$TestScript,

        [Parameter()]
        [bool]$Required = $true
    )

    try {
        & $TestScript | Out-Null

        return [PSCustomObject]@{
            Category   = "RBAC"
            Service    = $Service
            Test       = $TestName
            Status     = "Passed"
            Blocking   = $false
            RbacStatus = "Authorized"
            Details    = "The read operation completed successfully."
        }
    }
    catch {
        $message = $_.Exception.Message

        $authorizationPatterns = @(
            "Access is denied",
            "AccessDenied",
            "Unauthorized",
            "not authorized",
            "isn't within your assigned roles",
            "You don't have permission",
            "You do not have permission",
            "permission is required",
            "insufficient privileges"
        )

        $appearsRbacRelated = $false

        foreach ($pattern in $authorizationPatterns) {
            if ($message -match $pattern) {
                $appearsRbacRelated = $true
                break
            }
        }

        return [PSCustomObject]@{
            Category   = "RBAC"
            Service    = $Service
            Test       = $TestName
            Status     = "Failed"
            Blocking   = $Required
            RbacStatus = if ($appearsRbacRelated) {
                "Unauthorized"
            }
            else {
                "FailedOrUnavailable"
            }
            Details    = $message
        }
    }
}

function Invoke-PreflightValidation {
    [CmdletBinding()]
    param(
        [Parameter()]
        [bool]$ComplianceConnected = $false
    )

    Write-RetentionLog "Running the cmdlet and effective RBAC preflight."

    $results = [System.Collections.Generic.List[object]]::new()

    $exchangeCommands = @(
        @{
            Name     = "Get-Mailbox"
            Required = $true
        },
        @{
            Name     = "Get-OrganizationConfig"
            Required = $true
        },
        @{
            Name     = "Get-RetentionPolicy"
            Required = $true
        },
        @{
            Name     = "Get-RetentionPolicyTag"
            Required = $true
        }
    )

    foreach ($commandDefinition in $exchangeCommands) {
        $results.Add(
            (
                Test-CommandAvailability `
                    -CommandName $commandDefinition.Name `
                    -Service "Exchange Online" `
                    -Required $commandDefinition.Required
            )
        )
    }

    if (
        Get-Command `
            -Name Get-OrganizationConfig `
            -ErrorAction SilentlyContinue
    ) {
        $results.Add(
            (
                Test-ReadAccess `
                    -TestName "Read organization configuration" `
                    -Service "Exchange Online" `
                    -Required $true `
                    -TestScript {
                        Get-OrganizationConfig `
                            -ErrorAction Stop
                    }
            )
        )
    }

    if (
        Get-Command `
            -Name Get-Mailbox `
            -ErrorAction SilentlyContinue
    ) {
        $results.Add(
            (
                Test-ReadAccess `
                    -TestName "Read mailbox configuration" `
                    -Service "Exchange Online" `
                    -Required $true `
                    -TestScript {
                        Get-Mailbox `
                            -ResultSize 1 `
                            -ErrorAction Stop
                    }
            )
        )
    }

    if (
        Get-Command `
            -Name Get-RetentionPolicy `
            -ErrorAction SilentlyContinue
    ) {
        $results.Add(
            (
                Test-ReadAccess `
                    -TestName "Read MRM retention policies" `
                    -Service "Exchange Online" `
                    -Required $true `
                    -TestScript {
                        Get-RetentionPolicy `
                            -ErrorAction Stop
                    }
            )
        )
    }

    if (
        Get-Command `
            -Name Get-RetentionPolicyTag `
            -ErrorAction SilentlyContinue
    ) {
        $results.Add(
            (
                Test-ReadAccess `
                    -TestName "Read MRM retention tags" `
                    -Service "Exchange Online" `
                    -Required $true `
                    -TestScript {
                        Get-RetentionPolicyTag `
                            -ErrorAction Stop
                    }
            )
        )
    }

    if ($ComplianceConnected) {
        $purviewCommands = @(
            @{
                Name     = "Get-RetentionCompliancePolicy"
                Required = $true
            },
            @{
                Name     = "Get-RetentionComplianceRule"
                Required = $true
            },
            @{
                Name     = "Get-ComplianceCase"
                Required = $false
            },
            @{
                Name     = "Get-CaseHoldPolicy"
                Required = $false
            },
            @{
                Name     = "Get-AppRetentionCompliancePolicy"
                Required = $false
            }
        )

        foreach ($commandDefinition in $purviewCommands) {
            $results.Add(
                (
                    Test-CommandAvailability `
                        -CommandName $commandDefinition.Name `
                        -Service "Microsoft Purview" `
                        -Required $commandDefinition.Required
                )
            )
        }

        if (
            Get-Command `
                -Name Get-RetentionCompliancePolicy `
                -ErrorAction SilentlyContinue
        ) {
            $results.Add(
                (
                    Test-ReadAccess `
                        -TestName "Read Purview retention policies" `
                        -Service "Microsoft Purview" `
                        -Required $true `
                        -TestScript {
                            Get-RetentionCompliancePolicy `
                                -ErrorAction Stop
                        }
                )
            )
        }

        if (
            Get-Command `
                -Name Get-RetentionComplianceRule `
                -ErrorAction SilentlyContinue
        ) {
            $results.Add(
                (
                    Test-ReadAccess `
                        -TestName "Read Purview retention rules" `
                        -Service "Microsoft Purview" `
                        -Required $true `
                        -TestScript {
                            Get-RetentionComplianceRule `
                                -ErrorAction Stop
                        }
                )
            )
        }

        if (
            Get-Command `
                -Name Get-CaseHoldPolicy `
                -ErrorAction SilentlyContinue
        ) {
            if (
                Get-Command `
                    -Name Get-ComplianceCase `
                    -ErrorAction SilentlyContinue
            ) {
                $results.Add(
                    (
                        Test-ReadAccess `
                            -TestName "Read eDiscovery case hold policies" `
                            -Service "Microsoft Purview" `
                            -Required $false `
                            -TestScript {
                                Get-VisibleCaseHoldPolicy -ErrorAction Stop
                            }
                    )
                )
            }
        }

        if (
            Get-Command `
                -Name Get-AppRetentionCompliancePolicy `
                -ErrorAction SilentlyContinue
        ) {
            $results.Add(
                (
                    Test-ReadAccess `
                        -TestName "Read app retention policies" `
                        -Service "Microsoft Purview" `
                        -Required $false `
                        -TestScript {
                            Get-AppRetentionCompliancePolicy `
                                -ErrorAction Stop
                        }
                )
            )
        }
    }
    else {
        $results.Add(
            [PSCustomObject]@{
                Category   = "Connection"
                Service    = "Microsoft Purview"
                Test       = "Purview connection"
                Status     = "Skipped"
                Blocking   = $false
                RbacStatus = "NotTested"
                Details    = "The Microsoft Purview preflight was skipped because no compliance connection is available."
            }
        )
    }

    return $results
}

#endregion

#region Hold decoding

function Get-HoldActionDescription {
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowNull()]
        [string]$ActionCode
    )

    switch ($ActionCode) {
        "1" {
            return "Delete-only policy or policy associated with publishing or applying retention labels"
        }
        "2" {
            return "Retain-only policy"
        }
        "3" {
            return "Retain-and-delete policy"
        }
        default {
            return "Action not specified or not applicable"
        }
    }
}

function Resolve-InPlaceHoldValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$HoldValue,

        [Parameter()]
        [ValidateSet("Mailbox", "Organization")]
        [string]$Source = "Mailbox"
    )

    $originalValue = $HoldValue
    $workingValue = $HoldValue.Trim()

    $isExclusion = $false
    $prefix = $null
    $guidText = $null
    $actionCode = $null
    $holdType = "Unrecognized hold"
    $description = "The hold format was not recognized."
    $targetWorkload = "Undetermined"

    if ($workingValue.StartsWith("-")) {
        $isExclusion = $true
        $workingValue = $workingValue.Substring(1)
    }

    if (
        $workingValue -match
        "^(?<prefix>UniH)(?<guid>[0-9a-fA-F-]{32,36})$"
    ) {
        $prefix = $Matches.prefix
        $guidText = $Matches.guid
        $holdType = "eDiscovery hold"
        $targetWorkload = "Mailbox or workload included in an eDiscovery case"
        $description = "Unified eDiscovery hold"
    }
    elseif (
        $workingValue -match
        "^(?<prefix>mbx|skp|grp)(?<guid>[0-9a-fA-F-]{32,36})(:(?<action>[123]))?$"
    ) {
        $prefix = $Matches.prefix
        $guidText = $Matches.guid
        $actionCode = $Matches.action

        switch ($prefix) {
            "mbx" {
                $holdType = "Microsoft Purview retention policy"
                $targetWorkload = "Exchange mailbox or mailbox content"
            }
            "skp" {
                $holdType = "Microsoft Purview retention policy"
                $targetWorkload = "Skype for Business conversations stored in Exchange"
            }
            "grp" {
                $holdType = "Microsoft Purview retention policy"
                $targetWorkload = "Microsoft 365 Group or Teams content"
            }
        }

        if ($isExclusion -and $prefix -eq "mbx") {
            $holdType = "Organization-wide retention policy exclusion"
            $description = "The mailbox is explicitly excluded from the organization-wide policy."
        }
        else {
            $description = Get-HoldActionDescription `
                -ActionCode $actionCode
        }
    }
    elseif (
        $workingValue -match
        "^(?<guid>[0-9a-fA-F]{32}|[0-9a-fA-F-]{36})$"
    ) {
        $guidText = $Matches.guid
        $holdType = "Legacy In-Place Hold or unprefixed hold"
        $targetWorkload = "Exchange mailbox"
        $description = "Legacy or unprefixed hold value"
    }

    $normalizedGuid = $null

    if (-not [string]::IsNullOrWhiteSpace($guidText)) {
        $guidCandidate = $guidText

        if (
            $guidCandidate.Length -eq 32 -and
            $guidCandidate -notmatch "-"
        ) {
            $guidCandidate = "{0}-{1}-{2}-{3}-{4}" -f `
                $guidCandidate.Substring(0, 8),
                $guidCandidate.Substring(8, 4),
                $guidCandidate.Substring(12, 4),
                $guidCandidate.Substring(16, 4),
                $guidCandidate.Substring(20, 12)
        }

        try {
            $normalizedGuid = ([guid]$guidCandidate).Guid
        }
        catch {
            $normalizedGuid = $guidText
        }
    }

    return [PSCustomObject]@{
        OriginalValue      = $originalValue
        Source             = $Source
        IsExclusion        = $isExclusion
        Prefix             = $prefix
        HoldGuid           = $normalizedGuid
        ActionCode         = $actionCode
        ActionDescription  = Get-HoldActionDescription `
                                 -ActionCode $actionCode
        HoldType           = $holdType
        TargetWorkload     = $targetWorkload
        Description        = $description
        ResolvedObjectType = $null
        ResolvedName       = $null
        ResolvedStatus     = $null
        ResolutionError    = $null
    }
}

function Resolve-HoldObject {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$DecodedHold,

        [Parameter(Mandatory)]
        [hashtable]$RetentionPolicyCache,

        [Parameter(Mandatory)]
        [hashtable]$CaseHoldCache
    )

    if (
        [string]::IsNullOrWhiteSpace(
            [string]$DecodedHold.HoldGuid
        )
    ) {
        return $DecodedHold
    }

    $guidKey = (
        [string]$DecodedHold.HoldGuid
    ).ToLowerInvariant()

    if ($DecodedHold.Prefix -eq "UniH") {
        if ($CaseHoldCache.ContainsKey($guidKey)) {
            $caseHold = $CaseHoldCache[$guidKey]

            $DecodedHold.ResolvedObjectType =
                "eDiscovery Case Hold Policy"

            $DecodedHold.ResolvedName = Get-SafePropertyValue `
                -InputObject $caseHold `
                -PropertyName "Name"

            $DecodedHold.ResolvedStatus = Get-SafePropertyValue `
                -InputObject $caseHold `
                -PropertyName "Enabled"
        }
        else {
            $DecodedHold.ResolutionError =
                "The hold GUID was not found among the eDiscovery case hold policies visible to the connected account."
        }

        return $DecodedHold
    }

    if ($RetentionPolicyCache.ContainsKey($guidKey)) {
        $policy = $RetentionPolicyCache[$guidKey]

        $DecodedHold.ResolvedObjectType =
            "Retention Compliance Policy"

        $DecodedHold.ResolvedName = Get-SafePropertyValue `
            -InputObject $policy `
            -PropertyName "Name"

        $DecodedHold.ResolvedStatus = Get-SafePropertyValue `
            -InputObject $policy `
            -PropertyName "Enabled"
    }
    else {
        $DecodedHold.ResolutionError =
            "The hold GUID was not found among the retention policies visible to the connected account."
    }

    return $DecodedHold
}

#endregion

#region Initialization

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"

$reportPath = Join-Path `
    -Path $OutputPath `
    -ChildPath "Retention-Assessment-$timestamp"

New-Item `
    -Path $reportPath `
    -ItemType Directory `
    -Force |
    Out-Null

$transcriptStarted = $false
$exchangeConnected = $false
$complianceConnected = $false
$script:collectionFailures = $collectionFailures

try {
    Start-Transcript `
        -Path (Join-Path -Path $reportPath -ChildPath "Execution.log") `
        -Force |
        Out-Null

    $transcriptStarted = $true
}
catch {
    Write-Warning (
        "Unable to start the transcript: {0}" -f
        $_.Exception.Message
    )
}

Write-RetentionLog "Output directory: $reportPath"

#endregion

try {
    #region Runtime preflight

    $runtimePreflight = @(
        Test-RuntimeRequirement
    )

    Export-DataSet `
        -Name "00-Runtime-Preflight" `
        -Data $runtimePreflight `
        -Destination $reportPath

    $blockingRuntimeFailures = @(
        $runtimePreflight |
            Where-Object {
                $_.Status -eq "Failed" -and
                $_.Blocking -eq $true
            }
    )

    if ($blockingRuntimeFailures.Count -gt 0) {
        throw (
            "The runtime preflight failed: {0}" -f
            ($blockingRuntimeFailures.Test -join ", ")
        )
    }

    #endregion

    #region Connections

    Write-RetentionLog "Connecting to Exchange Online."

    if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        Connect-ExchangeOnline -ShowBanner:$false
    }
    else {
        Connect-ExchangeOnline `
            -UserPrincipalName $UserPrincipalName `
            -ShowBanner:$false
    }

    $exchangeConnected = $true

    if (-not $SkipComplianceConnection) {
        Write-RetentionLog "Connecting to Microsoft Purview PowerShell."

        try {
            if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
                Connect-IPPSSession -ShowBanner:$false
            }
            else {
                Connect-IPPSSession `
                    -UserPrincipalName $UserPrincipalName `
                    -ShowBanner:$false
            }

            $complianceConnected = $true
        }
        catch {
            $collectionFailures.Add(
                [PSCustomObject]@{
                    Section = "Microsoft Purview retention policies"
                    Details = $_.Exception.Message
                }
            )
            $collectionFailures.Add(
                [PSCustomObject]@{
                    Section = "Microsoft Purview connection"
                    Details = $_.Exception.Message
                }
            )
            Write-RetentionLog `
                -Level WARNING `
                -Message (
                    "The Microsoft Purview connection failed: {0}" -f
                    $_.Exception.Message
                )
        }
    }
    else {
        Write-RetentionLog `
            -Level WARNING `
            -Message "The Microsoft Purview connection was skipped by request."
    }

    #endregion

    #region Cmdlet and effective RBAC preflight

    $accessPreflight = @(
        Invoke-PreflightValidation `
            -ComplianceConnected $complianceConnected
    )

    Export-DataSet `
        -Name "01-Cmdlet-RBAC-Preflight" `
        -Data $accessPreflight `
        -Destination $reportPath

    $accessFailures = @(
        $accessPreflight |
            Where-Object {
                $_.Status -eq "Failed"
            }
    )

    $blockingAccessFailures = @(
        $accessFailures |
            Where-Object {
                $_.Blocking -eq $true
            }
    )

    if (
        $blockingAccessFailures.Count -gt 0 -and
        -not $ContinueOnPreflightFailure
    ) {
        $failureDescription = @(
            $blockingAccessFailures |
                ForEach-Object {
                    "{0}: {1}" -f $_.Test, $_.Details
                }
        ) -join " | "

        throw (
            "The cmdlet or effective RBAC preflight failed. {0}" -f
            $failureDescription
        )
    }

    if ($accessFailures.Count -gt 0) {
        Write-RetentionLog `
            -Level WARNING `
            -Message (
                "The assessment will continue with partial access. Preflight failures: {0}" -f
                ($accessFailures.Test -join ", ")
            )
    }

    #endregion

    #region Access flags

    $canReadOrganization = Test-PreflightPassed `
        -PreflightResults $accessPreflight `
        -TestName "Read organization configuration"

    $canReadMailboxes = Test-PreflightPassed `
        -PreflightResults $accessPreflight `
        -TestName "Read mailbox configuration"

    $canReadMrmPolicies = Test-PreflightPassed `
        -PreflightResults $accessPreflight `
        -TestName "Read MRM retention policies"

    $canReadMrmTags = Test-PreflightPassed `
        -PreflightResults $accessPreflight `
        -TestName "Read MRM retention tags"

    $canReadPurviewPolicies = Test-PreflightPassed `
        -PreflightResults $accessPreflight `
        -TestName "Read Purview retention policies"

    $canReadPurviewRules = Test-PreflightPassed `
        -PreflightResults $accessPreflight `
        -TestName "Read Purview retention rules"

    $canReadCaseHoldPolicies = Test-PreflightPassed `
        -PreflightResults $accessPreflight `
        -TestName "Read eDiscovery case hold policies"

    $canReadAppRetentionPolicies = Test-PreflightPassed `
        -PreflightResults $accessPreflight `
        -TestName "Read app retention policies"

    $collectionFailures = [System.Collections.Generic.List[object]]::new()
    $script:collectionFailures = $collectionFailures

    #endregion

    #region Organization configuration

    $organizationConfig = $null
    $organizationReport = @()
    $organizationHoldValues = @()
    $organizationHoldReport = @()

    if ($canReadOrganization) {
        Write-RetentionLog "Collecting the Exchange Online organization configuration."
        try {
            $organizationConfig = Get-OrganizationConfig `
                -ErrorAction Stop

            $organizationHoldValues = @(
                Get-SafePropertyValue `
                    -InputObject $organizationConfig `
                    -PropertyName "InPlaceHolds" `
                    -DefaultValue @()
            )

            $organizationReport = @(
                [PSCustomObject]@{
                    GeneratedAt = Get-Date

                    Name = Get-SafePropertyValue `
                        -InputObject $organizationConfig `
                        -PropertyName "Name"

                    Guid = Get-SafePropertyValue `
                        -InputObject $organizationConfig `
                        -PropertyName "Guid"

                    InPlaceHolds = Convert-ToFlatString `
                        -Value $organizationHoldValues

                    InPlaceHoldsCount = $organizationHoldValues.Count

                    ElcProcessingDisabled = Get-SafePropertyValue `
                        -InputObject $organizationConfig `
                        -PropertyName "ElcProcessingDisabled"

                    ComplianceMLIngestionEnabled = Get-SafePropertyValue `
                        -InputObject $organizationConfig `
                        -PropertyName "ComplianceMLIngestionEnabled"
                }
            )
        }
        catch {
            $collectionFailures.Add(
                [PSCustomObject]@{
                    Section = "Microsoft Purview retention policies"
                    Details = $_.Exception.Message
                }
            )
            Write-RetentionLog `
                -Level WARNING `
                -Message (
                    "Unable to retrieve Exchange Online organization configuration: {0}" -f
                    $_.Exception.Message
                )
        }
    }
    else {
        Write-RetentionLog `
            -Level WARNING `
            -Message "The organization configuration section was skipped because read access is unavailable."
    }

    #endregion

    #region Microsoft Purview configuration

    $retentionCompliancePolicies = @()
    $retentionComplianceRules = @()
    $caseHoldPolicies = @()
    $appRetentionPolicies = @()

    if ($canReadPurviewPolicies) {
        Write-RetentionLog "Collecting Microsoft Purview retention policies."

        try {
            $retentionCompliancePolicies = @(
                Get-RetentionCompliancePolicy `
                    -DistributionDetail `
                    -ErrorAction Stop
            )
        }
        catch {
            $collectionFailures.Add(
                [PSCustomObject]@{
                    Section = "Organization configuration"
                    Details = $_.Exception.Message
                }
            )
            Write-RetentionLog `
                -Level WARNING `
                -Message (
                    "Unable to retrieve Microsoft Purview retention policies: {0}" -f
                    $_.Exception.Message
                )
        }
    }

    if ($canReadPurviewRules) {
        Write-RetentionLog "Collecting Microsoft Purview retention rules."

        try {
            $retentionComplianceRules = @(
                Get-RetentionComplianceRule `
                    -ErrorAction Stop
            )
        }
        catch {
            $collectionFailures.Add(
                [PSCustomObject]@{
                    Section = "Microsoft Purview retention rules"
                    Details = $_.Exception.Message
                }
            )
            Write-RetentionLog `
                -Level WARNING `
                -Message (
                    "Unable to retrieve Microsoft Purview retention rules: {0}" -f
                    $_.Exception.Message
                )
        }
    }

    if ($canReadCaseHoldPolicies) {
        Write-RetentionLog "Collecting eDiscovery case hold policies."

        try {
            $caseHoldPolicies = @(
                Get-VisibleCaseHoldPolicy -RecordFailures -ErrorAction Stop
            )
        }
        catch {
            $collectionFailures.Add(
                [PSCustomObject]@{
                    Section = "eDiscovery case hold policies"
                    Details = $_.Exception.Message
                }
            )
            Write-RetentionLog `
                -Level WARNING `
                -Message (
                    "Unable to retrieve eDiscovery case hold policies: {0}" -f
                    $_.Exception.Message
                )
        }
    }

    if ($canReadAppRetentionPolicies) {
        Write-RetentionLog "Collecting app retention policies."

        try {
            $appRetentionPolicies = @(
                Get-AppRetentionCompliancePolicy `
                    -DistributionDetail `
                    -ErrorAction Stop
            )
        }
        catch {
            $collectionFailures.Add(
                [PSCustomObject]@{
                    Section = "App retention policies"
                    Details = $_.Exception.Message
                }
            )
            Write-RetentionLog `
                -Level WARNING `
                -Message (
                    "Unable to retrieve app retention policies: {0}" -f
                    $_.Exception.Message
                )
        }
    }

    $retentionPolicyReport = @(foreach (
        $policy in $retentionCompliancePolicies
    ) {
        [PSCustomObject]@{
            Name = Get-SafePropertyValue `
                -InputObject $policy `
                -PropertyName "Name"

            Guid = Get-SafePropertyValue `
                -InputObject $policy `
                -PropertyName "Guid"

            Enabled = Get-SafePropertyValue `
                -InputObject $policy `
                -PropertyName "Enabled"

            Mode = Get-SafePropertyValue `
                -InputObject $policy `
                -PropertyName "Mode"

            DistributionStatus = Get-SafePropertyValue `
                -InputObject $policy `
                -PropertyName "DistributionStatus"

            RestrictiveRetention = Get-SafePropertyValue `
                -InputObject $policy `
                -PropertyName "RestrictiveRetention"

            ExchangeLocation = Convert-ToFlatString `
                -Value (
                    Get-SafePropertyValue `
                        -InputObject $policy `
                        -PropertyName "ExchangeLocation" `
                        -DefaultValue @()
                )

            ExchangeLocationException = Convert-ToFlatString `
                -Value (
                    Get-SafePropertyValue `
                        -InputObject $policy `
                        -PropertyName "ExchangeLocationException" `
                        -DefaultValue @()
                )

            ModernGroupLocation = Convert-ToFlatString `
                -Value (
                    Get-SafePropertyValue `
                        -InputObject $policy `
                        -PropertyName "ModernGroupLocation" `
                        -DefaultValue @()
                )

            ModernGroupLocationException = Convert-ToFlatString `
                -Value (
                    Get-SafePropertyValue `
                        -InputObject $policy `
                        -PropertyName "ModernGroupLocationException" `
                        -DefaultValue @()
                )

            SharePointLocation = Convert-ToFlatString `
                -Value (
                    Get-SafePropertyValue `
                        -InputObject $policy `
                        -PropertyName "SharePointLocation" `
                        -DefaultValue @()
                )

            SharePointLocationException = Convert-ToFlatString `
                -Value (
                    Get-SafePropertyValue `
                        -InputObject $policy `
                        -PropertyName "SharePointLocationException" `
                        -DefaultValue @()
                )

            OneDriveLocation = Convert-ToFlatString `
                -Value (
                    Get-SafePropertyValue `
                        -InputObject $policy `
                        -PropertyName "OneDriveLocation" `
                        -DefaultValue @()
                )

            OneDriveLocationException = Convert-ToFlatString `
                -Value (
                    Get-SafePropertyValue `
                        -InputObject $policy `
                        -PropertyName "OneDriveLocationException" `
                        -DefaultValue @()
                )

            TeamsChatLocation = Convert-ToFlatString `
                -Value (
                    Get-SafePropertyValue `
                        -InputObject $policy `
                        -PropertyName "TeamsChatLocation" `
                        -DefaultValue @()
                )

            TeamsChatLocationException = Convert-ToFlatString `
                -Value (
                    Get-SafePropertyValue `
                        -InputObject $policy `
                        -PropertyName "TeamsChatLocationException" `
                        -DefaultValue @()
                )

            TeamsChannelLocation = Convert-ToFlatString `
                -Value (
                    Get-SafePropertyValue `
                        -InputObject $policy `
                        -PropertyName "TeamsChannelLocation" `
                        -DefaultValue @()
                )

            TeamsChannelLocationException = Convert-ToFlatString `
                -Value (
                    Get-SafePropertyValue `
                        -InputObject $policy `
                        -PropertyName "TeamsChannelLocationException" `
                        -DefaultValue @()
                )

            AdaptiveScopeLocation = Convert-ToFlatString `
                -Value (
                    Get-SafePropertyValue `
                        -InputObject $policy `
                        -PropertyName "AdaptiveScopeLocation" `
                        -DefaultValue @()
                )

            CreatedBy = Get-SafePropertyValue `
                -InputObject $policy `
                -PropertyName "CreatedBy"

            CreationDate = Get-SafePropertyValue `
                -InputObject $policy `
                -PropertyName "CreationDate"

            LastModifiedBy = Get-SafePropertyValue `
                -InputObject $policy `
                -PropertyName "LastModifiedBy"

            ModificationDate = Get-SafePropertyValue `
                -InputObject $policy `
                -PropertyName "ModificationDate"

            Comment = Get-SafePropertyValue `
                -InputObject $policy `
                -PropertyName "Comment"
        }
    })

    $retentionRuleReport = @(foreach (
        $rule in $retentionComplianceRules
    ) {
        [PSCustomObject]@{
            Name = Get-SafePropertyValue `
                -InputObject $rule `
                -PropertyName "Name"

            Guid = Get-SafePropertyValue `
                -InputObject $rule `
                -PropertyName "Guid"

            Policy = Get-SafePropertyValue `
                -InputObject $rule `
                -PropertyName "Policy"

            Disabled = Get-SafePropertyValue `
                -InputObject $rule `
                -PropertyName "Disabled"

            RetentionDuration = Get-SafePropertyValue `
                -InputObject $rule `
                -PropertyName "RetentionDuration"

            RetentionDurationDisplayHint = Get-SafePropertyValue `
                -InputObject $rule `
                -PropertyName "RetentionDurationDisplayHint"

            ExpirationDateOption = Get-SafePropertyValue `
                -InputObject $rule `
                -PropertyName "ExpirationDateOption"

            RetentionComplianceAction = Get-SafePropertyValue `
                -InputObject $rule `
                -PropertyName "RetentionComplianceAction"

            ApplyComplianceTag = Get-SafePropertyValue `
                -InputObject $rule `
                -PropertyName "ApplyComplianceTag"

            ContentMatchQuery = Get-SafePropertyValue `
                -InputObject $rule `
                -PropertyName "ContentMatchQuery"

            Workload = Convert-ToFlatString `
                -Value (
                    Get-SafePropertyValue `
                        -InputObject $rule `
                        -PropertyName "Workload" `
                        -DefaultValue @()
                )

            CreatedBy = Get-SafePropertyValue `
                -InputObject $rule `
                -PropertyName "CreatedBy"

            CreationDate = Get-SafePropertyValue `
                -InputObject $rule `
                -PropertyName "CreationDate"

            LastModifiedBy = Get-SafePropertyValue `
                -InputObject $rule `
                -PropertyName "LastModifiedBy"

            ModificationDate = Get-SafePropertyValue `
                -InputObject $rule `
                -PropertyName "ModificationDate"

            Comment = Get-SafePropertyValue `
                -InputObject $rule `
                -PropertyName "Comment"
        }
    })

    #endregion

    #region Exchange MRM configuration

    $mrmPolicies = @()
    $mrmTags = @()

    if ($canReadMrmPolicies) {
        Write-RetentionLog "Collecting Exchange MRM retention policies."
        try {
            $mrmPolicies = @(
                Get-RetentionPolicy -ErrorAction Stop
            )
        }
        catch {
            $collectionFailures.Add(
                [PSCustomObject]@{
                    Section = "Exchange MRM retention policies"
                    Details = $_.Exception.Message
                }
            )
            Write-RetentionLog `
                -Level WARNING `
                -Message (
                    "Unable to retrieve Exchange MRM retention policies: {0}" -f
                    $_.Exception.Message
                )
        }
    }

    if ($canReadMrmTags) {
        Write-RetentionLog "Collecting Exchange MRM retention tags."
        try {
            $mrmTags = @(
                Get-RetentionPolicyTag -ErrorAction Stop
            )
        }
        catch {
            $collectionFailures.Add(
                [PSCustomObject]@{
                    Section = "Exchange MRM retention tags"
                    Details = $_.Exception.Message
                }
            )
            Write-RetentionLog `
                -Level WARNING `
                -Message (
                    "Unable to retrieve Exchange MRM retention tags: {0}" -f
                    $_.Exception.Message
                )
        }
    }

    $mrmPolicyReport = @(foreach ($policy in $mrmPolicies) {
        [PSCustomObject]@{
            Name = Get-SafePropertyValue `
                -InputObject $policy `
                -PropertyName "Name"

            Guid = Get-SafePropertyValue `
                -InputObject $policy `
                -PropertyName "Guid"

            IsDefault = Get-SafePropertyValue `
                -InputObject $policy `
                -PropertyName "IsDefault"

            RetentionPolicyTagLinks = Convert-ToFlatString `
                -Value (
                    Get-SafePropertyValue `
                        -InputObject $policy `
                        -PropertyName "RetentionPolicyTagLinks" `
                        -DefaultValue @()
                )

            AdminDisplayName = Get-SafePropertyValue `
                -InputObject $policy `
                -PropertyName "AdminDisplayName"

            WhenCreated = Get-SafePropertyValue `
                -InputObject $policy `
                -PropertyName "WhenCreated"

            WhenChanged = Get-SafePropertyValue `
                -InputObject $policy `
                -PropertyName "WhenChanged"
        }
    })

    $mrmTagReport = @(foreach ($tag in $mrmTags) {
        [PSCustomObject]@{
            Name = Get-SafePropertyValue `
                -InputObject $tag `
                -PropertyName "Name"

            Guid = Get-SafePropertyValue `
                -InputObject $tag `
                -PropertyName "Guid"

            RetentionId = Get-SafePropertyValue `
                -InputObject $tag `
                -PropertyName "RetentionId"

            Type = Get-SafePropertyValue `
                -InputObject $tag `
                -PropertyName "Type"

            RetentionEnabled = Get-SafePropertyValue `
                -InputObject $tag `
                -PropertyName "RetentionEnabled"

            RetentionAction = Get-SafePropertyValue `
                -InputObject $tag `
                -PropertyName "RetentionAction"

            AgeLimitForRetention = Get-SafePropertyValue `
                -InputObject $tag `
                -PropertyName "AgeLimitForRetention"

            MessageClass = Get-SafePropertyValue `
                -InputObject $tag `
                -PropertyName "MessageClass"

            SystemTag = Get-SafePropertyValue `
                -InputObject $tag `
                -PropertyName "SystemTag"

            MustDisplayCommentEnabled = Get-SafePropertyValue `
                -InputObject $tag `
                -PropertyName "MustDisplayCommentEnabled"

            Comment = Get-SafePropertyValue `
                -InputObject $tag `
                -PropertyName "Comment"

            WhenCreated = Get-SafePropertyValue `
                -InputObject $tag `
                -PropertyName "WhenCreated"

            WhenChanged = Get-SafePropertyValue `
                -InputObject $tag `
                -PropertyName "WhenChanged"
        }
    })

    #endregion

    #region Hold resolution caches

    $retentionPolicyCache = @{}

    foreach ($policy in $retentionCompliancePolicies) {
        $policyGuid = Get-SafePropertyValue `
            -InputObject $policy `
            -PropertyName "Guid"

        if ($null -ne $policyGuid) {
            $key = ([string]$policyGuid).ToLowerInvariant()
            $retentionPolicyCache[$key] = $policy
        }
    }

    $caseHoldCache = @{}

    foreach ($caseHold in $caseHoldPolicies) {
        $caseHoldGuid = Get-SafePropertyValue `
            -InputObject $caseHold `
            -PropertyName "Guid"

        if ($null -ne $caseHoldGuid) {
            $key = ([string]$caseHoldGuid).ToLowerInvariant()
            $caseHoldCache[$key] = $caseHold
        }
    }

    #endregion

    #region Organization-wide hold decoding

    if ($canReadOrganization) {
        Write-RetentionLog "Decoding organization-wide holds."

        $organizationHoldReport = @(
            foreach ($holdValue in $organizationHoldValues) {
                if (
                    [string]::IsNullOrWhiteSpace(
                        [string]$holdValue
                    )
                ) {
                    continue
                }

                $decodedHold = Resolve-InPlaceHoldValue `
                    -HoldValue ([string]$holdValue) `
                    -Source "Organization"

                if ($complianceConnected) {
                    $decodedHold = Resolve-HoldObject `
                        -DecodedHold $decodedHold `
                        -RetentionPolicyCache $retentionPolicyCache `
                        -CaseHoldCache $caseHoldCache
                }

                $decodedHold
            }
        )
    }

    #endregion

    #region Mailbox inventory

    $mailboxReport = [System.Collections.Generic.List[object]]::new()
    $mailboxHoldReport = [System.Collections.Generic.List[object]]::new()
    $mailboxes = @()

    if ($canReadMailboxes) {
        Write-RetentionLog "Collecting mailboxes."
        try {
            $mailboxes = @(
                Get-Mailbox `
                    -ResultSize Unlimited `
                    -ErrorAction Stop
            )
        }
        catch {
            $collectionFailures.Add(
                [PSCustomObject]@{
                    Section = "Mailboxes"
                    Details = $_.Exception.Message
                }
            )
            Write-RetentionLog `
                -Level WARNING `
                -Message (
                    "Unable to retrieve mailboxes: {0}" -f
                    $_.Exception.Message
                )
        }

        if ($IncludeSoftDeletedMailboxes) {
            try {
                $softDeletedMailboxes = @(
                    Get-Mailbox `
                        -SoftDeletedMailbox `
                        -ResultSize Unlimited `
                        -ErrorAction Stop
                )

                $mailboxes += $softDeletedMailboxes
            }
            catch {
                $collectionFailures.Add(
                    [PSCustomObject]@{
                        Section = "Soft-deleted mailboxes"
                        Details = $_.Exception.Message
                    }
                )
                Write-RetentionLog `
                    -Level WARNING `
                    -Message (
                        "Unable to retrieve soft-deleted mailboxes: {0}" -f
                        $_.Exception.Message
                    )
            }
        }

        Write-RetentionLog "Mailboxes returned: $($mailboxes.Count)"

        $counter = 0

        foreach ($mailbox in $mailboxes) {
            $counter++

            $primarySmtpAddress = Get-SafePropertyValue `
                -InputObject $mailbox `
                -PropertyName "PrimarySmtpAddress"

            if ($mailboxes.Count -gt 0) {
                Write-Progress `
                    -Activity "Analyzing mailbox retention settings" `
                    -Status (
                        "{0} of {1}: {2}" -f
                        $counter,
                        $mailboxes.Count,
                        $primarySmtpAddress
                    ) `
                    -PercentComplete (
                        ($counter / $mailboxes.Count) * 100
                    )
            }

            $mailboxHolds = @(
                Get-SafePropertyValue `
                    -InputObject $mailbox `
                    -PropertyName "InPlaceHolds" `
                    -DefaultValue @()
            )

            $excludedPolicyGuids = @(
                foreach ($holdValue in $mailboxHolds) {
                    if (
                        [string]::IsNullOrWhiteSpace(
                            [string]$holdValue
                        )
                    ) {
                        continue
                    }

                    $decodedExclusion = Resolve-InPlaceHoldValue `
                        -HoldValue ([string]$holdValue) `
                        -Source "Mailbox"

                    if (
                        $decodedExclusion.IsExclusion -and
                        -not [string]::IsNullOrWhiteSpace(
                            [string]$decodedExclusion.HoldGuid
                        )
                    ) {
                        $decodedExclusion.HoldGuid
                    }
                }
            )

            $effectiveOrganizationHolds = @(
                foreach ($organizationHold in $organizationHoldReport) {
                    if (
                        -not [string]::IsNullOrWhiteSpace(
                            [string]$organizationHold.HoldGuid
                        ) -and
                        $excludedPolicyGuids -contains
                        $organizationHold.HoldGuid
                    ) {
                        continue
                    }

                    $organizationHold.OriginalValue
                }
            )

            $mrmPolicyName = Get-SafePropertyValue `
                -InputObject $mailbox `
                -PropertyName "RetentionPolicy"

            $mailboxPolicyTags = @()

            if (
                $canReadMrmPolicies -and
                -not [string]::IsNullOrWhiteSpace(
                    [string]$mrmPolicyName
                )
            ) {
                $policyObject = $mrmPolicies |
                    Where-Object {
                        (
                            Get-SafePropertyValue `
                                -InputObject $_ `
                                -PropertyName "Name"
                        ) -eq $mrmPolicyName
                    } |
                    Select-Object -First 1

                if ($null -ne $policyObject) {
                    $mailboxPolicyTags = @(
                        Get-SafePropertyValue `
                            -InputObject $policyObject `
                            -PropertyName "RetentionPolicyTagLinks" `
                            -DefaultValue @()
                    )
                }
            }

            $mailboxReport.Add(
                [PSCustomObject]@{
                    DisplayName = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "DisplayName"

                    PrimarySmtpAddress = $primarySmtpAddress

                    UserPrincipalName = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "UserPrincipalName"

                    ExternalDirectoryObjectId = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "ExternalDirectoryObjectId"

                    ExchangeGuid = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "ExchangeGuid"

                    RecipientTypeDetails = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "RecipientTypeDetails"

                    IsInactiveMailbox = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "IsInactiveMailbox"

                    IsSoftDeletedByRemove = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "IsSoftDeletedByRemove"

                    AccountDisabled = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "AccountDisabled"

                    MRMRetentionPolicy = $mrmPolicyName

                    MRMRetentionPolicyTags = Convert-ToFlatString `
                        -Value $mailboxPolicyTags

                    RetentionHoldEnabled = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "RetentionHoldEnabled"

                    RetentionComment = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "RetentionComment"

                    RetentionUrl = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "RetentionUrl"

                    StartDateForRetentionHold = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "StartDateForRetentionHold"

                    EndDateForRetentionHold = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "EndDateForRetentionHold"

                    ElcProcessingDisabled = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "ElcProcessingDisabled"

                    LitigationHoldEnabled = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "LitigationHoldEnabled"

                    LitigationHoldDuration = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "LitigationHoldDuration"

                    LitigationHoldDate = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "LitigationHoldDate"

                    LitigationHoldOwner = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "LitigationHoldOwner"

                    InPlaceHolds = Convert-ToFlatString `
                        -Value $mailboxHolds

                    MailboxInPlaceHoldsCount = $mailboxHolds.Count

                    OrganizationWideHolds = Convert-ToFlatString `
                        -Value $organizationHoldValues

                    EffectiveOrganizationWideHolds = Convert-ToFlatString `
                        -Value $effectiveOrganizationHolds

                    EffectiveOrganizationHoldsCount =
                        $effectiveOrganizationHolds.Count

                    ExcludedOrganizationPolicyGuids = Convert-ToFlatString `
                        -Value $excludedPolicyGuids

                    ComplianceTagHoldApplied = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "ComplianceTagHoldApplied"

                    DelayHoldApplied = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "DelayHoldApplied"

                    DelayReleaseHoldApplied = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "DelayReleaseHoldApplied"

                    SingleItemRecoveryEnabled = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "SingleItemRecoveryEnabled"

                    RetainDeletedItemsFor = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "RetainDeletedItemsFor"

                    ArchiveStatus = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "ArchiveStatus"

                    ArchiveState = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "ArchiveState"

                    ArchiveGuid = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "ArchiveGuid"

                    AutoExpandingArchiveEnabled = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "AutoExpandingArchiveEnabled"

                    AuditEnabled = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "AuditEnabled"

                    WhenCreated = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "WhenCreated"

                    WhenChanged = Get-SafePropertyValue `
                        -InputObject $mailbox `
                        -PropertyName "WhenChanged"
                }
            )

            foreach ($holdValue in $mailboxHolds) {
                if (
                    [string]::IsNullOrWhiteSpace(
                        [string]$holdValue
                    )
                ) {
                    continue
                }

                $decodedHold = Resolve-InPlaceHoldValue `
                    -HoldValue ([string]$holdValue) `
                    -Source "Mailbox"

                if ($complianceConnected) {
                    $decodedHold = Resolve-HoldObject `
                        -DecodedHold $decodedHold `
                        -RetentionPolicyCache $retentionPolicyCache `
                        -CaseHoldCache $caseHoldCache
                }

                $mailboxHoldReport.Add(
                    [PSCustomObject]@{
                        DisplayName = Get-SafePropertyValue `
                            -InputObject $mailbox `
                            -PropertyName "DisplayName"

                        PrimarySmtpAddress = $primarySmtpAddress

                        RecipientType = Get-SafePropertyValue `
                            -InputObject $mailbox `
                            -PropertyName "RecipientTypeDetails"

                        OriginalValue = $decodedHold.OriginalValue
                        IsExclusion = $decodedHold.IsExclusion
                        Prefix = $decodedHold.Prefix
                        HoldGuid = $decodedHold.HoldGuid
                        ActionCode = $decodedHold.ActionCode
                        ActionDescription = $decodedHold.ActionDescription
                        HoldType = $decodedHold.HoldType
                        TargetWorkload = $decodedHold.TargetWorkload
                        Description = $decodedHold.Description
                        ResolvedObjectType = $decodedHold.ResolvedObjectType
                        ResolvedName = $decodedHold.ResolvedName
                        ResolvedStatus = $decodedHold.ResolvedStatus
                        ResolutionError = $decodedHold.ResolutionError
                    }
                )
            }
        }

        Write-Progress `
            -Activity "Analyzing mailbox retention settings" `
            -Completed
    }
    else {
        Write-RetentionLog `
            -Level WARNING `
            -Message "The mailbox assessment section was skipped because read access is unavailable."
    }

    #endregion

    #region Summary

    $loadedModule = Get-Module `
        -Name ExchangeOnlineManagement |
        Sort-Object Version -Descending |
        Select-Object -First 1

    $summary = [PSCustomObject]@{
        GeneratedAt = Get-Date

        AssessmentStatus = if (
            $accessFailures.Count -eq 0 -and
            $collectionFailures.Count -eq 0
        ) {
            "Completed"
        }
        else {
            "CompletedWithPartialAccess"
        }

        PowerShellVersion =
            $PSVersionTable.PSVersion.ToString()

        ExchangeOnlineManagementVersion = if (
            $null -ne $loadedModule
        ) {
            $loadedModule.Version.ToString()
        }
        else {
            $null
        }

        ExchangeConnected = $exchangeConnected
        PurviewConnected = $complianceConnected

        PreflightFailureCount = @(
            $accessPreflight |
                Where-Object {
                    $_.Status -eq "Failed"
                }
        ).Count

        BlockingPreflightFailureCount =
            $blockingAccessFailures.Count

        CollectionFailureCount =
            $collectionFailures.Count

        MailboxCount = $mailboxReport.Count

        MailboxesWithLitigationHold = @(
            $mailboxReport |
                Where-Object {
                    $_.LitigationHoldEnabled -eq $true
                }
        ).Count

        MailboxesWithInPlaceHolds = @(
            $mailboxReport |
                Where-Object {
                    $_.MailboxInPlaceHoldsCount -gt 0
                }
        ).Count

        MailboxesWithRetentionHold = @(
            $mailboxReport |
                Where-Object {
                    $_.RetentionHoldEnabled -eq $true
                }
        ).Count

        MailboxesWithDelayHold = @(
            $mailboxReport |
                Where-Object {
                    $_.DelayHoldApplied -eq $true
                }
        ).Count

        MailboxesWithDelayReleaseHold = @(
            $mailboxReport |
                Where-Object {
                    $_.DelayReleaseHoldApplied -eq $true
                }
        ).Count

        MailboxesWithComplianceTagHold = @(
            $mailboxReport |
                Where-Object {
                    $_.ComplianceTagHoldApplied -eq $true
                }
        ).Count

        MailboxesWithElcProcessingDisabled = @(
            $mailboxReport |
                Where-Object {
                    $_.ElcProcessingDisabled -eq $true
                }
        ).Count

        MailboxesWithMRMPolicy = @(
            $mailboxReport |
                Where-Object {
                    -not [string]::IsNullOrWhiteSpace(
                        [string]$_.MRMRetentionPolicy
                    )
                }
        ).Count

        OrganizationWideHoldCount =
            $organizationHoldReport.Count

        RetentionCompliancePolicyCount =
            $retentionPolicyReport.Count

        RetentionComplianceRuleCount =
            $retentionRuleReport.Count

        MRMRetentionPolicyCount =
            $mrmPolicyReport.Count

        MRMRetentionTagCount =
            $mrmTagReport.Count

        CaseHoldPolicyCount =
            $caseHoldPolicies.Count

        AppRetentionPolicyCount =
            $appRetentionPolicies.Count
    }

    #endregion

    #region Export

    Write-RetentionLog "Exporting assessment reports."

    Export-DataSet `
        -Name "02-Summary" `
        -Data $summary `
        -Destination $reportPath

    Export-DataSet `
        -Name "03-OrganizationConfig" `
        -Data $organizationReport `
        -Destination $reportPath

    Export-DataSet `
        -Name "04-OrganizationWideHolds-Decoded" `
        -Data $organizationHoldReport `
        -Destination $reportPath

    Export-DataSet `
        -Name "05-Purview-RetentionPolicies" `
        -Data $retentionPolicyReport `
        -Destination $reportPath

    Export-DataSet `
        -Name "06-Purview-RetentionRules" `
        -Data $retentionRuleReport `
        -Destination $reportPath

    Export-DataSet `
        -Name "07-eDiscovery-CaseHoldPolicies" `
        -Data $caseHoldPolicies `
        -Destination $reportPath

    Export-DataSet `
        -Name "08-AppRetentionPolicies" `
        -Data $appRetentionPolicies `
        -Destination $reportPath

    Export-DataSet `
        -Name "09-Exchange-MRMPolicies" `
        -Data $mrmPolicyReport `
        -Destination $reportPath

    Export-DataSet `
        -Name "10-Exchange-MRMTags" `
        -Data $mrmTagReport `
        -Destination $reportPath

    Export-DataSet `
        -Name "11-Mailbox-RetentionConfiguration" `
        -Data $mailboxReport `
        -Destination $reportPath

    Export-DataSet `
        -Name "12-Mailbox-InPlaceHolds-Decoded" `
        -Data $mailboxHoldReport `
        -Destination $reportPath

    Export-DataSet `
        -Name "13-CollectionFailures" `
        -Data $collectionFailures `
        -Destination $reportPath

    Write-RetentionLog `
        -Level $(if ($summary.AssessmentStatus -eq "Completed") {
            "SUCCESS"
        }
        else {
            "WARNING"
        }) `
        -Message (
            "The assessment finished with status '{0}'. Reports were saved to: {1}" -f
            $summary.AssessmentStatus,
            $reportPath
        )

    $summary | Format-List

    #endregion
}
catch {
    Write-RetentionLog `
        -Level ERROR `
        -Message $_.Exception.Message

    throw
}
finally {
    if ($exchangeConnected) {
        try {
            Disconnect-ExchangeOnline `
                -Confirm:$false `
                -ErrorAction SilentlyContinue
        }
        catch {
            Write-Warning (
                "The Exchange Online disconnection returned an error: {0}" -f
                $_.Exception.Message
            )
        }
    }

    if ($transcriptStarted) {
        try {
            Stop-Transcript `
                -ErrorAction SilentlyContinue |
                Out-Null
        }
        catch {
            Write-Warning (
                'Unable to stop the transcript: {0}' -f
                $_.Exception.Message
            )
        }
    }
}