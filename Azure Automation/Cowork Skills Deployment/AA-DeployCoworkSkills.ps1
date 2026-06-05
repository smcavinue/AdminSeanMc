param(
    [string]$TenantName,
    [string]$SiteName
)

Function CopyFileToOneDrive {
    param(
        [string]$TenantName,
        [string]$SiteName,
        [string]$TargetUserName,
        [string]$SourceFilePath
    )

    $TargetOneDrive = get-mguserdrive -UserId sean.mcavinue@seanmcavinue.net | ? { $_.Name -eq "OneDrive" }
    $TargetFilePath = "/drives/$($TargetOneDrive.Id)/root:/Documents/Cowork/$($SourceFilePath.Split('\')[-1])"
    copy-mgdriveitem -DriveId $TargetOneDrive.Id -ItemId "root:/Documents/Cowork/$($SourceFilePath.Split('\')[-1])" -SourceFilePath $SourceFilePath
    #Write-Host "Copying file to $($TargetUserName)'s OneDrive at $TargetFilePath"
    #Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0$TargetFilePath:/copy" -Body @{
    #    parentReference = @{
    #        driveId = $OneDriveSite.Id
    #    }
    #    name = $SourceFilePath.Split('\')[-1]
    #} | Out-Null





}

Connect-mggraph #-Identity

$CoworkSite = get-mgsite -SiteId "$tenantname.sharepoint.com:/sites/$SiteName" 

$Documents = get-mgsitedrive -SiteId $CoworkSite.Id  -Filter "Name eq 'Documents'"

$List = get-mgsiteList -SiteId $CoworkSite.Id -Filter "displayName eq 'Documents'"

[array]$items = Get-MgSiteListItem -SiteId $CoworkSite.Id -ListId $List.id -Select 'fields' -ExpandProperty 'fields'


$Items = Invoke-MgGraphRequest -method Get -Uri "https://graph.microsoft.com/v1.0/sites/$($Coworksite.Id)/lists/$($List.Id)/items?`$expand=fields(`$select=id,Title,DistributeTo,DistributeToLookupId)"

$Deployments = $Items.value | Where-Object { $_.fields.DistributeTo -Ne $null }

$PeopleList = get-mgsiteList -SiteId $CoworkSite.Id -Filter "displayName eq 'User Information List'"

foreach ($Deployment in $Deployments) {

    $file = Invoke-MgGraphRequest -method Get -Uri "https://graph.microsoft.com/v1.0/sites/$($Coworksite.Id)/lists/$($List.Id)/items/$($deployment.id)/driveitem"

    $User = Invoke-MgGraphRequest -method Get -Uri "https://graph.microsoft.com/v1.0/sites/$($Coworksite.Id)/lists/$($PeopleList.Id)/items/$($Deployment.fields.DistributeToLookupId)?expand=fields"
    
    If ($Deployment.fields.DistributeTo -eq "Everyone except external users") {
        write-host "Deployed to EEEU Group"
        
    }
    elseif ($User.fields.ContentType -eq "Person") {
        write-host "Deployed to User: $($User.fields.Title)"
    }
    elseIf ($User.fields.ContentType -eq "DomainGroup") {
        write-host "Deployed to Group: $($Deployment.fields.DistributeTo)"
    }
}


