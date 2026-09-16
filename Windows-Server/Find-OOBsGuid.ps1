param(
    [string]$OutputPath = (Join-Path $PSScriptRoot "OOB-GUIDs.csv")
)

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

# Optional export
$Results | Export-Csv $OutputPath -NoTypeInformation