##Declare Variables
$AdminURL = "<SharePointAdminURL>"
$SiteID = "<SharePointSiteID>"
$ListID = "<SharePointListID>"

##Connect to Microsoft Graph
Connect-MgGraph -Identity -NoWelcome

##Connect to PNP Online
Connect-PnPOnline -ManagedIdentity -Url $adminURL

##Get all sites that are marked for deletion
[array]$Sites = Get-MgSiteListItem -SiteId $SiteID -ListId $listID -Filter "fields/MarkedforDeletion eq 'Yes' and fields/Deletedon eq null"  -Expand "fields(`$select=SiteURL)" -All -Select id, fields

if(!$sites){
    write-output "No sites to remove"
}

$Fields = @{
    fields = @{    
        Deletedon = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
    }

}

##Delete each site
ForEach ($Site in $Sites) {
write-output "Deleting: $($site.Fields.AdditionalProperties.SiteURL)"
    ##Check if there is an associated group
    $siteRelativeURL = $site.Fields.AdditionalProperties.SiteURL.replace("https://", "").replace("/sites/", ":/sites/")
    $SiteObject = Get-MgSite -SiteId $siteRelativeURL
    $GroupObject = get-mggroup -Filter "displayName eq '$($SiteObject.displayName)'" -ErrorAction silentlycontinue

    ##If a group exists, delete the group, otherwise delete the site
    if ($GroupObject) {
        ##Delete the group
        write-output "Site is a group connected site: $($site.Fields.AdditionalProperties.SiteURL)"
        Remove-MgGroup -GroupId $GroupObject.id -Confirm:$false
    }
    else {
        ##Delete site
        write-output "Site is a communication site: $($site.Fields.AdditionalProperties.SiteURL)"
        Remove-PnPTenantSite -Url $site.Fields.AdditionalProperties.SiteURL -Force
    }

    ##Update list item to reflect that the site has been deleted
    Update-MgSiteListItem -SiteId $SiteID -ListId $ListID -ListItemId $Site.Id -BodyParameter $Fields
}

##Disconnect from PNP Online
Disconnect-PnPOnline
##Disconnect from Microsoft Graph
Disconnect-MgGraph