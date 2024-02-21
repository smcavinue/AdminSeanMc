$TenantAdminURL = "https://<TenantID>-admin.sharepoint.com"
$Account = "<AdminAccount>@<TenantID>.onmicrosoft.com"
$LabelID = "LabelID"
# Connect to SharePoint Online
Connect-PnPOnline -Url $TenantAdminURL -Interactive

# Get all sites
$sites = Get-PnPTenantSite | ?{($_.template -like "Group#0") -or ($_.template -like "TeamChannel*")}

# Iterate through each site
foreach ($site in $sites) {
    # Add account as owner of the site temporarily
    Set-PnPTenantSite -Url $site.Url -Owners $Account
}

foreach ($site in $sites) {
    # Connect to the site
    Connect-PnPOnline -Url $site.Url -Interactive

    # Get all document libraries
    [array]$Libraries = Get-PnPList | Where-Object {$_.BaseTemplate -eq 101 -and $_.Hidden -eq $false -and $_.Title -eq "Shared Documents"}

    # Iterate through each document library
    foreach ($library in $libraries) {
        # Set the default sensitivity label
        Set-PnPList -Identity $libraries.Title -DefaultSensitivityLabelForLibrary $LabelID
    }

}

# Disconnect from SharePoint Online
Disconnect-PnPOnline
