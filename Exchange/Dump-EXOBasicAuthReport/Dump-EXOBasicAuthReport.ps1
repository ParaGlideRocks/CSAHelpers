#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Comprehensive Exchange Online Basic Authentication Report Script
    
.DESCRIPTION
    This script generates a detailed report of Exchange Online basic authentication settings including:
    - Organization-level authentication settings
    - All authentication policies and their configurations
    - Users with authentication policies that have basic auth enabled for any protocol
    - Default authentication policy assignment
    - Summary statistics
    
.PARAMETER OutputPath
    Path where the report files will be saved. Defaults to current directory.
    
.PARAMETER ExportFormat
    Export format for the report. Options: CSV, JSON, XML. Defaults to CSV.
    
.PARAMETER IncludeDetailedUserInfo
    Include additional user details like last sign-in, creation date, etc.
    
.EXAMPLE
    .\Get-EXOBasicAuthReport.ps1
    
.EXAMPLE
    .\Get-EXOBasicAuthReport.ps1 -OutputPath "C:\Reports" -ExportFormat JSON -IncludeDetailedUserInfo
    
.NOTES
    Author: Generated for Exchange Online Basic Auth Reporting
    Requires: Exchange Online PowerShell V3 module
    Permissions: Exchange Administrator or Global Reader role minimum
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".",
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("CSV", "JSON", "XML")]
    [string]$ExportFormat = "CSV",
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeDetailedUserInfo
)

# Initialize variables
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$reportFiles = @{}

Write-Host "Starting Exchange Online Basic Authentication Report..." -ForegroundColor Green
Write-Host "Timestamp: $(Get-Date)" -ForegroundColor Gray

try {
    # Check if connected to Exchange Online
    Write-Host "Checking Exchange Online connection..." -ForegroundColor Yellow
    
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        Write-Host "✓ Connected to Exchange Online" -ForegroundColor Green
    }
    catch {
        Write-Error "Not connected to Exchange Online. Please run Connect-ExchangeOnline first."
        return
    }

    # ====================================
    # 1. Organization Configuration Report
    # ====================================
    Write-Host "`n1. Gathering Organization Configuration..." -ForegroundColor Yellow
    
    $orgConfig = Get-OrganizationConfig | Select-Object -Property Name, DefaultAuthenticationPolicy
    $modernAuthConfig = Get-OrganizationConfig | Select-Object -Property OAuth2ClientProfileEnabled, MapiHttpEnabled
    
    # Check if modern authentication is enabled
    $authConfigDetails = [PSCustomObject]@{
        OrganizationName = $orgConfig.Name
        DefaultAuthenticationPolicy = if ($orgConfig.DefaultAuthenticationPolicy) { $orgConfig.DefaultAuthenticationPolicy } else { "None Set" }
        OAuth2ClientProfileEnabled = $modernAuthConfig.OAuth2ClientProfileEnabled
        MapiHttpEnabled = $modernAuthConfig.MapiHttpEnabled
        ReportGeneratedDate = Get-Date
        ModernAuthStatus = if ($modernAuthConfig.OAuth2ClientProfileEnabled) { "Enabled" } else { "Disabled" }
    }
    
    Write-Host "✓ Organization configuration collected" -ForegroundColor Green

    # ====================================
    # 2. Authentication Policies Report
    # ====================================
    Write-Host "`n2. Gathering Authentication Policies..." -ForegroundColor Yellow
    
    $authPolicies = Get-AuthenticationPolicy
    $authPolicyDetails = @()
    
    if ($authPolicies) {
        foreach ($policy in $authPolicies) {
            $policyDetail = [PSCustomObject]@{
                PolicyName = $policy.Name
                DistinguishedName = $policy.DistinguishedName
                AllowBasicAuthActiveSync = $policy.AllowBasicAuthActiveSync
                AllowBasicAuthAutodiscover = $policy.AllowBasicAuthAutodiscover
                AllowBasicAuthImap = $policy.AllowBasicAuthImap
                AllowBasicAuthMapi = $policy.AllowBasicAuthMapi
                AllowBasicAuthOfflineAddressBook = $policy.AllowBasicAuthOfflineAddressBook
                AllowBasicAuthOutlookService = $policy.AllowBasicAuthOutlookService
                AllowBasicAuthPop = $policy.AllowBasicAuthPop
                AllowBasicAuthReportingWebServices = $policy.AllowBasicAuthReportingWebServices
                AllowBasicAuthRpc = $policy.AllowBasicAuthRpc
                AllowBasicAuthSmtp = $policy.AllowBasicAuthSmtp
                AllowBasicAuthWebServices = $policy.AllowBasicAuthWebServices
                AllowBasicAuthPowershell = $policy.AllowBasicAuthPowershell
                IsDefault = ($policy.Name -eq $orgConfig.DefaultAuthenticationPolicy)
                HasAnyBasicAuthEnabled = ($policy.AllowBasicAuthActiveSync -or 
                                        $policy.AllowBasicAuthAutodiscover -or 
                                        $policy.AllowBasicAuthImap -or 
                                        $policy.AllowBasicAuthMapi -or 
                                        $policy.AllowBasicAuthOfflineAddressBook -or 
                                        $policy.AllowBasicAuthOutlookService -or 
                                        $policy.AllowBasicAuthPop -or 
                                        $policy.AllowBasicAuthReportingWebServices -or 
                                        $policy.AllowBasicAuthRpc -or 
                                        $policy.AllowBasicAuthSmtp -or 
                                        $policy.AllowBasicAuthWebServices -or 
                                        $policy.AllowBasicAuthPowershell)
            }
            $authPolicyDetails += $policyDetail
        }
        Write-Host "✓ Found $($authPolicies.Count) authentication policies" -ForegroundColor Green
    }
    else {
        Write-Host "! No authentication policies found" -ForegroundColor Yellow
        $authPolicyDetails = @([PSCustomObject]@{
            PolicyName = "No policies found"
            DistinguishedName = "N/A"
            AllowBasicAuthActiveSync = "N/A"
            AllowBasicAuthAutodiscover = "N/A"
            AllowBasicAuthImap = "N/A"
            AllowBasicAuthMapi = "N/A"
            AllowBasicAuthOfflineAddressBook = "N/A"
            AllowBasicAuthOutlookService = "N/A"
            AllowBasicAuthPop = "N/A"
            AllowBasicAuthReportingWebServices = "N/A"
            AllowBasicAuthRpc = "N/A"
            AllowBasicAuthSmtp = "N/A"
            AllowBasicAuthWebServices = "N/A"
            AllowBasicAuthPowershell = "N/A"
            IsDefault = $false
            HasAnyBasicAuthEnabled = $false
        })
    }

    # ====================================
    # 3. Users with Basic Auth Enabled
    # ====================================
    Write-Host "`n3. Gathering Users with Basic Authentication Policies..." -ForegroundColor Yellow
    
    $usersWithBasicAuth = @()
    $allUsers = @()
    
    # Get all users with authentication policies
    Write-Host "   Retrieving all users (this may take several minutes)..." -ForegroundColor Gray
    $allUsers = Get-User -ResultSize Unlimited | Where-Object { $_.AuthenticationPolicy -ne $null -or $orgConfig.DefaultAuthenticationPolicy -ne $null }
    
    Write-Host "   Analyzing $($allUsers.Count) users..." -ForegroundColor Gray
    
    foreach ($user in $allUsers) {
        $userAuthPolicy = $null
        $hasBasicAuthEnabled = $false
        $enabledProtocols = @()
        
        # Determine which policy applies to the user
        if ($user.AuthenticationPolicy) {
            $userAuthPolicy = $authPolicyDetails | Where-Object { $_.DistinguishedName -eq $user.AuthenticationPolicy }
        }
        elseif ($orgConfig.DefaultAuthenticationPolicy) {
            $userAuthPolicy = $authPolicyDetails | Where-Object { $_.PolicyName -eq $orgConfig.DefaultAuthenticationPolicy }
        }
        
        if ($userAuthPolicy) {
            # Check each protocol for basic auth enablement
            $protocolChecks = @{
                'ActiveSync' = $userAuthPolicy.AllowBasicAuthActiveSync
                'Autodiscover' = $userAuthPolicy.AllowBasicAuthAutodiscover
                'IMAP' = $userAuthPolicy.AllowBasicAuthImap
                'MAPI' = $userAuthPolicy.AllowBasicAuthMapi
                'OfflineAddressBook' = $userAuthPolicy.AllowBasicAuthOfflineAddressBook
                'OutlookService' = $userAuthPolicy.AllowBasicAuthOutlookService
                'POP' = $userAuthPolicy.AllowBasicAuthPop
                'ReportingWebServices' = $userAuthPolicy.AllowBasicAuthReportingWebServices
                'RPC' = $userAuthPolicy.AllowBasicAuthRpc
                'SMTP' = $userAuthPolicy.AllowBasicAuthSmtp
                'WebServices' = $userAuthPolicy.AllowBasicAuthWebServices
                'PowerShell' = $userAuthPolicy.AllowBasicAuthPowershell
            }
            
            foreach ($protocol in $protocolChecks.Keys) {
                if ($protocolChecks[$protocol] -eq $true) {
                    $hasBasicAuthEnabled = $true
                    $enabledProtocols += $protocol
                }
            }
            
            # Only include users who have basic auth enabled for at least one protocol
            if ($hasBasicAuthEnabled) {
                $userDetails = [PSCustomObject]@{
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName = $user.DisplayName
                    SamAccountName = $user.SamAccountName
                    RecipientType = $user.RecipientType
                    AuthenticationPolicy = if ($user.AuthenticationPolicy) { 
                        ($authPolicyDetails | Where-Object { $_.DistinguishedName -eq $user.AuthenticationPolicy }).PolicyName 
                    } else { 
                        "Default Policy: $($orgConfig.DefaultAuthenticationPolicy)" 
                    }
                    PolicyAssignmentType = if ($user.AuthenticationPolicy) { "Direct" } else { "Default" }
                    EnabledBasicAuthProtocols = ($enabledProtocols -join ', ')
                    ProtocolCount = $enabledProtocols.Count
                }
                
                # Add detailed user info if requested
                if ($IncludeDetailedUserInfo) {
                    try {
                        $mailbox = Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
                        $userDetails | Add-Member -NotePropertyName "CreatedDate" -NotePropertyValue $user.WhenCreated
                        $userDetails | Add-Member -NotePropertyName "LastModified" -NotePropertyValue $user.WhenChanged
                        $userDetails | Add-Member -NotePropertyName "MailboxType" -NotePropertyValue $mailbox.RecipientTypeDetails
                        $userDetails | Add-Member -NotePropertyName "IsLicensed" -NotePropertyValue ($user.AssignedLicenses.Count -gt 0)
                    }
                    catch {
                        # Continue if additional details can't be retrieved
                    }
                }
                
                $usersWithBasicAuth += $userDetails
            }
        }
    }
    
    Write-Host "✓ Found $($usersWithBasicAuth.Count) users with basic authentication enabled" -ForegroundColor Green

    # ====================================
    # 4. Generate Summary Statistics
    # ====================================
    Write-Host "`n4. Generating Summary Statistics..." -ForegroundColor Yellow
    
    $summary = [PSCustomObject]@{
        ReportDate = Get-Date
        TotalAuthenticationPolicies = $authPolicyDetails.Count
        PoliciesWithBasicAuthEnabled = ($authPolicyDetails | Where-Object { $_.HasAnyBasicAuthEnabled -eq $true }).Count
        TotalUsersAnalyzed = $allUsers.Count
        UsersWithBasicAuthEnabled = $usersWithBasicAuth.Count
        DefaultPolicyConfigured = ($orgConfig.DefaultAuthenticationPolicy -ne $null)
        DefaultPolicyName = $orgConfig.DefaultAuthenticationPolicy
        ModernAuthenticationEnabled = $modernAuthConfig.OAuth2ClientProfileEnabled
        
        # Protocol-specific statistics
        UsersWithActiveSync = ($usersWithBasicAuth | Where-Object { $_.EnabledBasicAuthProtocols -like "*ActiveSync*" }).Count
        UsersWithPOP = ($usersWithBasicAuth | Where-Object { $_.EnabledBasicAuthProtocols -like "*POP*" }).Count
        UsersWithIMAP = ($usersWithBasicAuth | Where-Object { $_.EnabledBasicAuthProtocols -like "*IMAP*" }).Count
        UsersWithSMTP = ($usersWithBasicAuth | Where-Object { $_.EnabledBasicAuthProtocols -like "*SMTP*" }).Count
        UsersWithMAPI = ($usersWithBasicAuth | Where-Object { $_.EnabledBasicAuthProtocols -like "*MAPI*" }).Count
    }
    
    Write-Host "✓ Summary statistics generated" -ForegroundColor Green

    # ====================================
    # 5. Export Reports
    # ====================================
    Write-Host "`n5. Exporting Reports..." -ForegroundColor Yellow
    
    $baseFileName = "EXO_BasicAuth_Report_$timestamp"
    
    # Export Organization Configuration
    $orgFileName = "$OutputPath\$($baseFileName)_OrgConfig.$($ExportFormat.ToLower())"
    switch ($ExportFormat) {
        "CSV" { $authConfigDetails | Export-Csv -Path $orgFileName -NoTypeInformation }
        "JSON" { $authConfigDetails | ConvertTo-Json -Depth 3 | Out-File -FilePath $orgFileName }
        "XML" { $authConfigDetails | Export-Clixml -Path $orgFileName }
    }
    $reportFiles["Organization Configuration"] = $orgFileName
    
    # Export Authentication Policies
    $policiesFileName = "$OutputPath\$($baseFileName)_AuthPolicies.$($ExportFormat.ToLower())"
    switch ($ExportFormat) {
        "CSV" { $authPolicyDetails | Export-Csv -Path $policiesFileName -NoTypeInformation }
        "JSON" { $authPolicyDetails | ConvertTo-Json -Depth 3 | Out-File -FilePath $policiesFileName }
        "XML" { $authPolicyDetails | Export-Clixml -Path $policiesFileName }
    }
    $reportFiles["Authentication Policies"] = $policiesFileName
    
    # Export Users with Basic Auth
    if ($usersWithBasicAuth.Count -gt 0) {
        $usersFileName = "$OutputPath\$($baseFileName)_UsersWithBasicAuth.$($ExportFormat.ToLower())"
        switch ($ExportFormat) {
            "CSV" { $usersWithBasicAuth | Export-Csv -Path $usersFileName -NoTypeInformation }
            "JSON" { $usersWithBasicAuth | ConvertTo-Json -Depth 3 | Out-File -FilePath $usersFileName }
            "XML" { $usersWithBasicAuth | Export-Clixml -Path $usersFileName }
        }
        $reportFiles["Users with Basic Auth"] = $usersFileName
    }
    
    # Export Summary
    $summaryFileName = "$OutputPath\$($baseFileName)_Summary.$($ExportFormat.ToLower())"
    switch ($ExportFormat) {
        "CSV" { $summary | Export-Csv -Path $summaryFileName -NoTypeInformation }
        "JSON" { $summary | ConvertTo-Json -Depth 3 | Out-File -FilePath $summaryFileName }
        "XML" { $summary | Export-Clixml -Path $summaryFileName }
    }
    $reportFiles["Summary"] = $summaryFileName
    
    Write-Host "✓ All reports exported successfully" -ForegroundColor Green

    # ====================================
    # 6. Display Summary to Console
    # ====================================
    Write-Host "`n" + "="*60 -ForegroundColor Cyan
    Write-Host "EXCHANGE ONLINE BASIC AUTHENTICATION REPORT SUMMARY" -ForegroundColor Cyan
    Write-Host "="*60 -ForegroundColor Cyan
    
    Write-Host "`nOrganization Overview:" -ForegroundColor White
    Write-Host "  Organization: $($authConfigDetails.OrganizationName)"
    Write-Host "  Modern Authentication: $($authConfigDetails.ModernAuthStatus)"
    Write-Host "  Default Auth Policy: $($authConfigDetails.DefaultAuthenticationPolicy)"
    
    Write-Host "`nAuthentication Policies:" -ForegroundColor White
    Write-Host "  Total Policies: $($summary.TotalAuthenticationPolicies)"
    Write-Host "  Policies with Basic Auth Enabled: $($summary.PoliciesWithBasicAuthEnabled)"
    
    Write-Host "`nUser Analysis:" -ForegroundColor White
    Write-Host "  Total Users Analyzed: $($summary.TotalUsersAnalyzed)"
    Write-Host "  Users with Basic Auth Enabled: $($summary.UsersWithBasicAuthEnabled)" 
    
    if ($summary.UsersWithBasicAuthEnabled -gt 0) {
        Write-Host "`nProtocol Breakdown:" -ForegroundColor White
        Write-Host "  ActiveSync: $($summary.UsersWithActiveSync) users"
        Write-Host "  POP3: $($summary.UsersWithPOP) users"
        Write-Host "  IMAP4: $($summary.UsersWithIMAP) users"
        Write-Host "  SMTP: $($summary.UsersWithSMTP) users"
        Write-Host "  MAPI: $($summary.UsersWithMAPI) users"
    }
    
    Write-Host "`nExported Files:" -ForegroundColor White
    foreach ($report in $reportFiles.Keys) {
        Write-Host "  $report`: $($reportFiles[$report])"
    }
    
    Write-Host "`n" + "="*60 -ForegroundColor Cyan
    
    # Security recommendations
    if ($summary.UsersWithBasicAuthEnabled -gt 0) {
        Write-Host "`nSECURITY RECOMMENDATIONS:" -ForegroundColor Red
        Write-Host "• Review users with basic authentication enabled"
        Write-Host "• Ensure clients support modern authentication before disabling basic auth"
        Write-Host "• Consider implementing Conditional Access policies"
        Write-Host "• Monitor sign-in logs for legacy authentication attempts"
        Write-Host "• Plan migration timeline for legacy clients"
    }
    else {
        Write-Host "`nSECURITY STATUS: GOOD" -ForegroundColor Green
        Write-Host "• No users found with basic authentication enabled"
        Write-Host "• Organization appears to be using modern authentication"
    }
    
    Write-Host "`nReport completed successfully at $(Get-Date)" -ForegroundColor Green
}
catch {
    Write-Error "An error occurred during report generation: $($_.Exception.Message)"
    Write-Error "Full error details: $($_.Exception)"
}

# ====================================
# Additional Helper Functions
# ====================================

<#
.SYNOPSIS
    Helper function to get detailed policy information for a specific user
    
.PARAMETER UserPrincipalName
    The UPN of the user to check
    
.EXAMPLE
    Get-UserBasicAuthStatus -UserPrincipalName "user@contoso.com"
#>
function Get-UserBasicAuthStatus {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName
    )
    
    try {
        $user = Get-User -Identity $UserPrincipalName
        $orgConfig = Get-OrganizationConfig
        
        $userPolicy = $null
        if ($user.AuthenticationPolicy) {
            $userPolicy = Get-AuthenticationPolicy -Identity $user.AuthenticationPolicy
        }
        elseif ($orgConfig.DefaultAuthenticationPolicy) {
            $userPolicy = Get-AuthenticationPolicy -Identity $orgConfig.DefaultAuthenticationPolicy
        }
        
        if ($userPolicy) {
            return [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                PolicyName = $userPolicy.Name
                PolicyType = if ($user.AuthenticationPolicy) { "Direct" } else { "Default" }
                ActiveSync = $userPolicy.AllowBasicAuthActiveSync
                Autodiscover = $userPolicy.AllowBasicAuthAutodiscover
                IMAP = $userPolicy.AllowBasicAuthImap
                MAPI = $userPolicy.AllowBasicAuthMapi
                OfflineAddressBook = $userPolicy.AllowBasicAuthOfflineAddressBook
                OutlookService = $userPolicy.AllowBasicAuthOutlookService
                POP = $userPolicy.AllowBasicAuthPop
                ReportingWebServices = $userPolicy.AllowBasicAuthReportingWebServices
                RPC = $userPolicy.AllowBasicAuthRpc
                SMTP = $userPolicy.AllowBasicAuthSmtp
                WebServices = $userPolicy.AllowBasicAuthWebServices
                PowerShell = $userPolicy.AllowBasicAuthPowershell
            }
        }
        else {
            return [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                PolicyName = "No Policy Applied"
                PolicyType = "None"
                Status = "All protocols default to organization settings"
            }
        }
    }
    catch {
        Write-Error "Error retrieving information for user $UserPrincipalName`: $($_.Exception.Message)"
    }
}

Write-Host "`nAdditional Functions Available:" -ForegroundColor Yellow
Write-Host "  Get-UserBasicAuthStatus -UserPrincipalName 'user@domain.com'" -ForegroundColor Gray
Write-Host "    Use this function to check specific user's basic auth status" -ForegroundColor Gray