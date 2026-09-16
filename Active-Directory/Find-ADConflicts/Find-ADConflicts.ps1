<#
.SYNOPSIS
    Identifies conflicts between AD Users and Contacts that may cause issues during Entra sync.
.DESCRIPTION
    Finds duplicate mail addresses, proxyAddresses, and display names between Users and Contacts.
#>
 
param(
    [string]$OutputPath = "$PSScriptRoot\ConflictReport.csv"
)
 
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
 
function ConvertTo-NormalizedAddress {
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )
 
    $trimmed = $Value.Trim()
    if (-not $trimmed) { return $null }
 
    # Handle proxyAddresses formats like SMTP:user@domain or smtp:user@domain
    $parts = $trimmed -split ':', 2
    if ($parts.Count -eq 2) {
        return $parts[1].Trim().ToLowerInvariant()
    }
 
    return $trimmed.ToLowerInvariant()
}
 
function New-ConflictRow {
    param(
        [Parameter(Mandatory)][string]$ConflictType,
        [Parameter(Mandatory)][string]$Severity,
        [Parameter(Mandatory)][object]$Contact,
        [Parameter(Mandatory)][object]$User,
        [string]$ContactMail,
        [string]$ContactProxyAddress,
        [string]$MatchedValue
    )
 
    [PSCustomObject]@{
        ConflictType        = $ConflictType
        Severity            = $Severity
        MatchedValue        = $MatchedValue
 
        ContactName         = $Contact.DisplayName
        ContactMail         = $ContactMail
        ContactProxyAddress = $ContactProxyAddress
 
        UserName            = $User.DisplayName
        UserMail            = $User.Mail
        UserPrincipal       = $User.UserPrincipalName
    }
}
 
$conflicts = New-Object System.Collections.Generic.List[object]
 
# Get all users and contacts
if (-not (Get-Command -Name Get-ADUser -ErrorAction SilentlyContinue)) {
    throw "Get-ADUser was not found. Install/Import the ActiveDirectory module (RSAT) and retry."
}
 
$users = Get-ADUser -Filter * -Properties mail, proxyAddresses, displayName, userPrincipalName
 
# Contacts are returned as ADObject; select a consistent property surface
$contacts = Get-ADObject -LDAPFilter '(objectClass=contact)' -Properties mail, proxyAddresses, displayName |
    Select-Object -Property @(
        @{ Name = 'DisplayName'; Expression = { $_.DisplayName } },
        @{ Name = 'Mail'; Expression = { $_.Mail } },
        @{ Name = 'ProxyAddresses'; Expression = { $_.ProxyAddresses } }
    )
 
## Build a unified address index so mail/proxy collisions are detected across both fields
$usersByAddress = @{}  # normalized address -> [System.Collections.Generic.List[object]]
 
foreach ($user in $users) {
    if ($user.Mail) {
        $key = ConvertTo-NormalizedAddress -Value $user.Mail
        if ($key) {
            if (-not $usersByAddress.ContainsKey($key)) {
                $usersByAddress[$key] = New-Object System.Collections.Generic.List[object]
            }
            $usersByAddress[$key].Add($user)
        }
    }
 
    foreach ($proxyAddr in @($user.ProxyAddresses)) {
        $key = if ($proxyAddr) { ConvertTo-NormalizedAddress -Value ([string]$proxyAddr) }
        if (-not $key) { continue }
 
        if (-not $usersByAddress.ContainsKey($key)) {
            $usersByAddress[$key] = New-Object System.Collections.Generic.List[object]
        }
        if (-not $usersByAddress[$key].Contains($user)) {
            $usersByAddress[$key].Add($user)
        }
    }
}
 
foreach ($contact in $contacts) {
    # Mail address conflicts
    if ($contact.Mail) {
        $contactMailKey = ConvertTo-NormalizedAddress -Value $contact.Mail
        if ($contactMailKey -and $usersByAddress.ContainsKey($contactMailKey)) {
            foreach ($user in $usersByAddress[$contactMailKey]) {
                $conflicts.Add((New-ConflictRow -ConflictType 'Mail Address' -Severity 'High' -Contact $contact -User $user -ContactMail $contact.Mail -MatchedValue $contactMailKey))
            }
        }
    }
 
    # proxyAddresses conflicts
    foreach ($proxyAddr in @($contact.ProxyAddresses)) {
        $proxyString = if ($proxyAddr) { [string]$proxyAddr }
        $proxyKey = if ($proxyString) { ConvertTo-NormalizedAddress -Value $proxyString }
        if (-not $proxyKey) { continue }
 
        if ($usersByAddress.ContainsKey($proxyKey)) {
            foreach ($user in $usersByAddress[$proxyKey]) {
                $conflicts.Add((New-ConflictRow -ConflictType 'ProxyAddress' -Severity 'Critical' -Contact $contact -User $user -ContactProxyAddress $proxyString -MatchedValue $proxyKey))
            }
        }
    }
}
 
# Export results
if ($conflicts.Count -gt 0) {
    $conflicts | Export-Csv -Path $OutputPath -NoTypeInformation
    Write-Host "Found $($conflicts.Count) conflicts. Report saved to: $OutputPath"
} else {
    Write-Host "No conflicts detected."
}