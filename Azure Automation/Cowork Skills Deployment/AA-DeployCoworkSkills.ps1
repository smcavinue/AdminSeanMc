param(
    [string]$TenantName,
    [string]$SiteName
)

Function CopyFileToOneDrive {
    param(
        [string]$TenantName,
        [string]$SiteName,
        [string]$TargetUserName,
        [string]$SourceFileId
    )

    $SourceFile = Get-MgDriveItem -DriveId $SourceDriveId -DriveItemId $SourceFileId

    $TargetOneDrive = get-mguserdrive -UserId sean.mcavinue@seanmcavinue.net | ? { $_.Name -eq "OneDrive" }
    $TargetFilePath = "/drives/$($TargetOneDrive.Id)/root:/Documents/Cowork/Skills"
    Try{
    $TargetCoworkFolder = get-mgdriveitem -DriveId $TargetOneDrive.Id -DriveItemId "root:/Documents/Cowork1"
    }catch{
        $DocumentsFolder = Get-MgDriveItem -DriveId $TargetOneDrive.Id -DriveItemId "root:/Documents"
        #Write-Host "Creating Cowork folder in $($TargetUserName)'s OneDrive"
        $TargetCoworkFolder = New-MgDriveItemChild -DriveId $TargetOneDrive.Id -Name "Cowork1" -Folder @{ childCount = 0 } -DriveItemId  $DocumentsFolder.Id
        $TargetSkillsFolder = New-MgDriveItemChild -DriveId $TargetOneDrive.Id -Name "Skills" -Folder @{ childCount = 0 } -DriveItemId  $TargetCoworkFolder.Id

    }

    Try{
    $TargetSkillsFolder = get-mgdriveitem -DriveId $TargetOneDrive.Id -DriveItemId "root:/Documents/Cowork1/Skills"
    }catch{
        $TargetSkillsFolder = New-MgDriveItemChild -DriveId $TargetOneDrive.Id -Name "Skills" -Folder @{ childCount = 0 } -DriveItemId  $TargetCoworkFolder.Id
    }

    $params = @{
        parentReference = @{
            driveId = $TargetOneDrive.Id
            id = $TargetSkillsFolder.Id
        }
        name = $SourceFile.Name
    }

    Copy-MgDriveItem -DriveId $SourceFile.ParentReference.DriveId -DriveItemId $SourceFile.Id -BodyParameter $params

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
    $sourceFilePath = $file.parentReference.path.Replace("/sites/$SiteName/drive/root:", "") + "/" + $file.name
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


