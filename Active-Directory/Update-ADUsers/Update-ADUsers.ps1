[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelPath,

    [string]$WorksheetName = "Users",

    [string]$LogPath = ".\AD_BulkUpdate_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",

    [switch]$StopOnError
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level, $Message

    switch ($Level) {
        "ERROR"   { Write-Host $line -ForegroundColor Red }
        "WARN"    { Write-Host $line -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $line -ForegroundColor Green }
        default   { Write-Host $line -ForegroundColor Cyan }
    }

    Add-Content -Path $LogPath -Value $line
}

function Throw-LoggedError {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Log -Level "ERROR" -Message $Message
    throw $Message
}

function Test-ClearValue {
    param([object]$Value)

    return $null -ne $Value -and $Value.ToString().Trim() -eq "__CLEAR__"
}

function Test-HasValue {
    param([object]$Value)

    if ($null -eq $Value) {
        return $false
    }

    if (Test-ClearValue -Value $Value) {
        return $true
    }

    return -not [string]::IsNullOrWhiteSpace($Value.ToString())
}

function Convert-ToNullableString {
    param([object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    $text = $Value.ToString().Trim()
    if ($text -eq "") {
        return $null
    }

    return $text
}

function Convert-ToBoolean {
    param([object]$Value)

    if ($null -eq $Value) {
        return $null
    }

    $text = $Value.ToString().Trim().ToLowerInvariant()

    switch ($text) {
        "true"   { return $true }
        "1"      { return $true }
        "yes"    { return $true }
        "y"      { return $true }
        "enabled" { return $true }

        "false"   { return $false }
        "0"       { return $false }
        "no"      { return $false }
        "n"       { return $false }
        "disabled" { return $false }

        default {
            Throw-LoggedError -Message "Invalid boolean value '$Value'. Use True/False, Yes/No, 1/0, Enabled/Disabled."
        }
    }
}

function Resolve-AdUser {
    param(
        [string]$SamAccountName,
        [string]$UserPrincipalName
    )

    if (-not [string]::IsNullOrWhiteSpace($SamAccountName)) {
        $adUser = @(Get-ADUser -Filter "SamAccountName -eq '$SamAccountName'" -Properties * -ErrorAction Stop)

        if ($adUser.Count -eq 0) {
            Write-Log -Level "WARN" -Message "Active Directory user with SamAccountName '$SamAccountName' was not found. Skipping this row."
            return $null
        }

        if ($adUser.Count -gt 1) {
            Write-Log -Level "WARN" -Message "Multiple Active Directory users were found with SamAccountName '$SamAccountName'. Skipping this row."
            return $null
        }

        return $adUser[0]
    }

    if (-not [string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        $adUser = @(Get-ADUser -Filter "UserPrincipalName -eq '$UserPrincipalName'" -Properties * -ErrorAction Stop)

        if ($adUser.Count -eq 0) {
            Write-Log -Level "WARN" -Message "Active Directory user with UserPrincipalName '$UserPrincipalName' was not found. Skipping this row."
            return $null
        }

        if ($adUser.Count -gt 1) {
            Write-Log -Level "WARN" -Message "Multiple Active Directory users were found with UserPrincipalName '$UserPrincipalName'. Skipping this row."
            return $null
        }

        return $adUser[0]
    }

    Throw-LoggedError -Message "Row does not contain SamAccountName or UserPrincipalName."
}

function Resolve-ManagerDn {
    param([string]$ManagerValue)

    if ([string]::IsNullOrWhiteSpace($ManagerValue)) {
        return $null
    }

    if ($ManagerValue -eq "__CLEAR__") {
        return "__CLEAR__"
    }

    try {
        $manager = Get-ADUser -Filter "SamAccountName -eq '$ManagerValue' -or UserPrincipalName -eq '$ManagerValue'" -Properties DistinguishedName -ErrorAction Stop
        return $manager.DistinguishedName
    }
    catch {
        Throw-LoggedError -Message "Manager '$ManagerValue' was not found in Active Directory."
    }
}

function Test-WorksheetExists {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $worksheetInfo = Get-ExcelSheetInfo -Path $Path -ErrorAction Stop
    return $worksheetInfo.Name -contains $Name
}

try {
    "" | Set-Content -Path $LogPath
    Write-Log -Message "Starting bulk AD update."
    Write-Log -Message "Excel file: $ExcelPath"
    Write-Log -Message "Worksheet: $WorksheetName"
    Write-Log -Message "Log file: $LogPath"

    if (-not (Test-Path -Path $ExcelPath)) {
        Throw-LoggedError -Message "Excel file not found: $ExcelPath"
    }

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Throw-LoggedError -Message "Module 'ImportExcel' is not installed. Install it with: Install-Module ImportExcel -Scope CurrentUser"
    }

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Throw-LoggedError -Message "Module 'ActiveDirectory' is not available. Install RSAT AD tools on this machine."
    }

    Import-Module ImportExcel -ErrorAction Stop
    Import-Module ActiveDirectory -ErrorAction Stop

    if ([string]::IsNullOrWhiteSpace($WorksheetName)) {
        Throw-LoggedError -Message "Worksheet name cannot be empty."
    }

    if (-not (Test-WorksheetExists -Path $ExcelPath -Name $WorksheetName)) {
        Throw-LoggedError -Message "Worksheet '$WorksheetName' was not found in Excel file '$ExcelPath'."
    }

    $rows = Import-Excel -Path $ExcelPath -WorksheetName $WorksheetName -ErrorAction Stop

    if (-not $rows -or $rows.Count -eq 0) {
        Throw-LoggedError -Message "No rows were found in worksheet '$WorksheetName'."
    }

    $processed = 0
    $updated = 0
    $failed = 0
    $skipped = 0

    foreach ($row in $rows) {
        $processed++

        try {
            $samAccountName = Convert-ToNullableString $row.SamAccountName
            $userPrincipalName = Convert-ToNullableString $row.UserPrincipalName

            Write-Log -Message "Processing row $processed. SamAccountName='$samAccountName', UPN='$userPrincipalName'"

            $adUser = Resolve-AdUser -SamAccountName $samAccountName -UserPrincipalName $userPrincipalName

            if ($null -eq $adUser) {
                $skipped++
                continue
            }

            $setParams = @{
                Identity    = $adUser.DistinguishedName
                ErrorAction = "Stop"
            }

            $clearAttributes = New-Object System.Collections.Generic.List[string]
            $replaceAttributes = @{}

            $directMap = @{
                GivenName     = "GivenName"
                Surname       = "Surname"
                DisplayName   = "DisplayName"
                Title         = "Title"
                Department    = "Department"
                Company       = "Company"
                Office        = "Office"
                TelephoneNumber   = "TelephoneNumber"
                Mobile   = "Mobile"
                Description   = "Description"
                StreetAddress = "StreetAddress"
                City          = "City"
                State         = "State"
                PostalCode    = "PostalCode"
                Country       = "Country"
            }

            foreach ($column in $directMap.Keys) {
                $value = $row.$column

                if (-not (Test-HasValue -Value $value)) {
                    continue
                }

                if (Test-ClearValue -Value $value) {
                    $ldapClearMap = @{
                        GivenName     = "givenName"
                        Surname       = "sn"
                        DisplayName   = "displayName"
                        EmailAddress  = "mail"
                        Title         = "title"
                        Department    = "department"
                        Company       = "company"
                        Office        = "physicalDeliveryOfficeName"
                        telephoneNumber   = "telephoneNumber"
                        Mobile   = "mobile"
                        Description   = "description"
                        StreetAddress = "streetAddress"
                        City          = "l"
                        State         = "st"
                        PostalCode    = "postalCode"
                        Country       = "co"
                    }

                    $clearAttributes.Add($ldapClearMap[$column])
                }
                else {
                    $setParams[$directMap[$column]] = $value.ToString().Trim()
                }
            }

            if (Test-HasValue -Value $row.Manager) {
                $managerDn = Resolve-ManagerDn -ManagerValue (Convert-ToNullableString $row.Manager)

                if ($managerDn -eq "__CLEAR__") {
                    $clearAttributes.Add("manager")
                }
                else {
                    $setParams["Manager"] = $managerDn
                }
            }

            for ($i = 1; $i -le 15; $i++) {
                $columnName = "ExtensionAttribute$i"
                $ldapName = "extensionAttribute$i"
                $value = $row.$columnName

                if (-not (Test-HasValue -Value $value)) {
                    continue
                }

                if (Test-ClearValue -Value $value) {
                    $clearAttributes.Add($ldapName)
                }
                else {
                    $replaceAttributes[$ldapName] = $value.ToString().Trim()
                }
            }

            if ($replaceAttributes.Count -gt 0) {
                $setParams["Replace"] = $replaceAttributes
            }

            if ($clearAttributes.Count -gt 0) {
                $setParams["Clear"] = $clearAttributes.ToArray()
            }

            $targetName = if ($adUser.SamAccountName) { $adUser.SamAccountName } else { $adUser.DistinguishedName }

            if ($PSCmdlet.ShouldProcess($targetName, "Update Active Directory properties")) {
                Set-ADUser @setParams

                if (Test-HasValue -Value $row.Enabled) {
                    $enabled = Convert-ToBoolean -Value $row.Enabled

                    if ($enabled) {
                        Enable-ADAccount -Identity $adUser.DistinguishedName -ErrorAction Stop
                    }
                    else {
                        Disable-ADAccount -Identity $adUser.DistinguishedName -ErrorAction Stop
                    }
                }
            }

            $updated++
            Write-Log -Level "SUCCESS" -Message "Updated user '$targetName'."
        }
        catch {
            $failed++
            Write-Log -Level "ERROR" -Message "Row $processed failed: $($_.Exception.Message)"

            if ($StopOnError) {
                Write-Log -Level "ERROR" -Message "StopOnError is enabled. Execution will stop after row $processed."
                throw
            }
        }
    }

    Write-Log -Message "Bulk AD update finished. Processed=$processed Updated=$updated Skipped=$skipped Failed=$failed"
}
catch {
    Write-Log -Level "ERROR" -Message "Execution stopped: $($_.Exception.Message)"
    return
}