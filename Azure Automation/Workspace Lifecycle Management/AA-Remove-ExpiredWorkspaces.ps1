##Declare Variables
$AdminURL = "<SharePointAdminURL>"
$SiteID = "<SharePointSiteID>"
$ListID = "<SharePointListID>"

##Connect to Microsoft Graph
Connect-MgGraph -Identity -NoWelcome

##Connect to PNP Online
Connect-PnPOnline -ManagedIdentity -Url $adminURL

##Get all sites that are marked for deletion
[array]$Sites = Get-MgSiteListItem -SiteId $SiteID -ListId $listID -Filter "fields/MarkedforDeletion eq 'Yes'"  -Expand "fields(`$select=SiteURL)" -All -Select id,fields

$Fields = @{
    fields = @{    
        Deletedon = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
    }

}

##Delete each site
ForEach ($Site in $Sites) {

    ##Delete site
    Remove-PnPTenantSite -Url $SiteURL -Force -confirm:$false

    ##Update list item to reflect that the site has been deleted
    Update-MgSiteListItem -SiteId $SiteID -ListId $ListID -ListItemId $Site.Id -BodyParameter $Fields
}

