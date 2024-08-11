param (
    [Parameter(Mandatory = $true)]
    [string]$CSVFilePath,
    [Parameter(Mandatory = $true)]
    [string]$SensitivityLabelID,
    [Parameter(Mandatory = $true)]
    [string]$Prefix,
    [Parameter(Mandatory = $true)]
    [string]$SharePointTenant
)

$Sites =  import-csv $CSVFilePath

$AdminURL = "https://$SharePointTenant"
Connect-PnPOnline -Url $AdminURL -Interactive

$createdSitesArray = @()

foreach ($site in $Sites) {
    $SiteName = "$Prefix - $($site.siteName)"

    $NewSite = New-PnPSite -Type TeamSite -Title $SiteName -SensitivityLabel $SensitivityLabelID -Alias $SiteName.replace(' ', '')
    
    $createdSitesArray += $NewSite
}

foreach ($createdSite in $createdSitesArray) {
    $siteURL = $createdSite
    $PageName = "Home.aspx"
    $DocumentLibraryName = "Documents"
    Connect-PnPOnline -Url $siteURL -Interactive

    $Page = get-pnppage -Identity $PageName
    $List = Get-PnpList -Identity $DocumentLibraryName
    Add-PnPPageSection -Page $Page -SectionTemplate OneColumn -Order 1
    Add-PnPPageTextPart -Page $Page -text "Welcome to your new site!" -Section 1 -Column 1
    Add-PnPPageWebPart -Page $page -DefaultWebPartType "List" -Section 1 -Column 1 -WebPartProperties @{ isDocumentLibrary = "true"; selectedListId = "$($List.id)" }
}