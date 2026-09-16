<#
.SYNOPSIS
Finds Microsoft Update Catalog GUIDs for a predefined list of out-of-band updates.

.DESCRIPTION
Searches the Microsoft Update Catalog for each KB number in the script, follows every
matching update detail link, and extracts the update GUID and associated products.
Results are displayed in the console and exported to a CSV file.

.PARAMETER OutputPath
Specifies the destination CSV file. The default is OOB-GUIDs.csv in the script directory.

.EXAMPLE
PS> .\Find-OOBsGuid.ps1

Displays the update results and writes them to OOB-GUIDs.csv.

.EXAMPLE
PS> .\Find-OOBsGuid.ps1 -OutputPath C:\Temp\OOB-GUIDs.csv

Writes the CSV results to the specified path.

.NOTES
Requires internet access to catalog.update.microsoft.com. The extraction logic depends
on the current HTML structure of the Microsoft Update Catalog pages.
#>
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot "OOB-GUIDs.csv")
)

# KB articles to locate in the Microsoft Update Catalog.
$KBs = @(
    "KB5129237",
    "KB5129235",
    "KB5129238",
    "KB5129239",
    "KB5129243",
    "KB5129244"
)

$Results = foreach ($KB in $KBs) {

    Write-Host "Searching $KB..."

    $SearchUrl = "https://www.catalog.update.microsoft.com/Search.aspx?q=$KB"

    try {
        $Response = Invoke-WebRequest -Uri $SearchUrl -UseBasicParsing

        # A search can return several catalog entries for different products or platforms.
        $DetailLinkMatches = [regex]::Matches(
            $Response.Content,
            'goToDetails\("(?<UpdateID>[0-9a-fA-F-]{36})"\)'
        )

        if ($DetailLinkMatches.Count -eq 0) {
            [PSCustomObject]@{
                KB       = $KB
                UpdateID = "NOT FOUND"
                Products = ""
            }
            continue
        }

        # Visit each result because the detail page contains the authoritative UpdateID
        # and the product list associated with that catalog entry.
        foreach ($DetailLinkMatch in $DetailLinkMatches) {
            $DetailsUrl = "https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid=$($DetailLinkMatch.Groups['UpdateID'].Value)"
            $DetailsResponse = Invoke-WebRequest -Uri $DetailsUrl -UseBasicParsing
            $UpdateIdMatch = [regex]::Match(
                $DetailsResponse.Content,
                'id="ScopedViewHandler_UpdateID">\s*(?<UpdateID>[0-9a-fA-F-]{36})'
            )

            if (-not $UpdateIdMatch.Success) {
                throw "The UpdateID was not found on the details page: $DetailsUrl"
            }

            $ProductsMatch = [regex]::Match(
                $DetailsResponse.Content,
                '<div id="productsDiv">.*?</span>(?<Products>.*?)</div>',
                [System.Text.RegularExpressions.RegexOptions]::Singleline
            )

            if (-not $ProductsMatch.Success) {
                throw "The Products field was not found on the details page: $DetailsUrl"
            }

            # Convert the products fragment into readable, whitespace-normalized text.
            $Products = [System.Net.WebUtility]::HtmlDecode(
                ([regex]::Replace($ProductsMatch.Groups['Products'].Value, '<[^>]+>', ' ') -replace '\s+', ' ').Trim()
            )

            [PSCustomObject]@{
                KB       = $KB
                UpdateID = $UpdateIdMatch.Groups['UpdateID'].Value
                Products = $Products
            }
        }
    }
    catch {
        [PSCustomObject]@{
            KB       = $KB
            UpdateID = "ERROR: $($_.Exception.Message)"
            Products = ""
        }
    }
}

$Results | Format-Table -AutoSize

# Persist the same objects shown in the console for later review or comparison.
$Results | Export-Csv $OutputPath -NoTypeInformation