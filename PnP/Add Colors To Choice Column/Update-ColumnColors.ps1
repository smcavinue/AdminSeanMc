$siteURL = Read-Host "Enter the site URL"
$clientID = Read-Host "Enter the client ID"
$tenantName = Read-Host "Enter the tenant name"

##Connect to the site
Connect-PnPOnline -Url $siteurl -Interactive -ClientId $clientID -Tenant $tenantName
##Get the Documents Library
$List = Get-PnPList -Identity "Documents"

Try {
    ##Try get the "Choice" column
    $field = Get-PnPField -List "Documents" -Identity "Choices"
}
catch {
    ##If not found write a warning
    Write-host "Choices field not found in $siteurl" -ForegroundColor yellow
    $field = $null
}
if ($field) {
    ##If found, import the formatter
    $Tags = $field.choices
    $Colors = import-csv .\JSONFormatter\Colors.csv
    $header = get-content .\JSONFormatter\Start-Formatter.json
    $Entry = get-content .\JSONFormatter\Per-item-Formatter.json
    $footer = get-content .\JSONFormatter\End-Formatter.json

    $Content = $header
    $x = 0

    ##Foreach choice, add the color to the formatter
    foreach ($tag in $tags) {
        if ($x -eq $colors.count) {
            $x = 0
        }
        $Content += ($Entry -replace "{TAGPLACEHOLDER}", $tag) -replace "{ColorPlaceholder}", $colors[$x].colors 
        $x++
    }

    $content = $content[0..($content.Length - 2)]
    $content += '"sp-field-borderAllRegular sp-field-borderAllSolid sp-css-borderColor-neutralSecondary"'

    foreach ($tag in $tags) {
        $Content += "]"
        $Content += "}"
    }

    ##Add the footer to the formatter
    $Content += $footer

    ##Save the formatter to a file
    Set-content -Value $Content -LiteralPath CustomFormatter.json -Encoding utf8

    ##Get the content of the formatter
    $JSON = get-content -Raw .\CustomFormatter.json

    ##Set the formatter to the column
    Set-PnPField -List $List.id -Identity $Field.id -Values @{CustomFormatter = $JSON.tostring() }

}
else {
    write-host "No Tags column present in $siteurl" -ForegroundColor yellow
}
