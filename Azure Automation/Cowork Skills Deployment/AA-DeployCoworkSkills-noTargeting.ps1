param(
    [string]$TenantName,
    [string]$SiteName,
    [Boolean]$OverwriteExisting = $false
)

Function CopyFileToOneDrive {
    param(
        [string]$TenantName,
        [string]$SiteName,
        [string]$SourceDriveId,
        [string]$TargetUserName,
        [string]$SourceFileId,
        [switch]$OverwriteExisting = $false
    )

    $SourceFile = Get-MgDriveItem -DriveId $SourceDriveId -DriveItemId $SourceFileId

    $TargetOneDrive = Get-MgUserDefaultDrive -UserId $TargetUserName 
    $TargetFilePath = "root:/Documents/Cowork/Skills"
    $TargetSkill = "root:/Documents/Cowork/Skills/$($SourceFile.Name)"


    If( (Get-MgDriveItem -DriveId $TargetOneDrive.Id -DriveItemId $TargetSkill -ErrorAction SilentlyContinue) -ne $null) {
        write-host "Skill already exists in target location ($TargetSkill), skipping deployment"
        If(-not $OverwriteExisting) {
            return
        }
        else {
            write-host "Overwriting existing skill in target location ($TargetSkill)"
            Remove-MgDriveItem -DriveId $TargetOneDrive.Id -DriveItemId $TargetSkill
        }
    }

    
    Try{
    $TargetCoworkFolder = get-mgdriveitem -DriveId $TargetOneDrive.Id -DriveItemId "root:/Documents/Cowork" -erroraction stop
    }catch{
        $DocumentsFolder = Get-MgDriveItem -DriveId $TargetOneDrive.Id -DriveItemId "root:/Documents"
        Write-Host "Creating Cowork folder in $($TargetUserName)'s OneDrive"
        $TargetCoworkFolder = New-MgDriveItemChild -DriveId $TargetOneDrive.Id -Name "Cowork" -Folder @{ childCount = 0 } -DriveItemId  $DocumentsFolder.Id
        $TargetSkillsFolder = New-MgDriveItemChild -DriveId $TargetOneDrive.Id -Name "Skills" -Folder @{ childCount = 0 } -DriveItemId  $TargetCoworkFolder.Id
    }

    Try{
    $TargetSkillsFolder = get-mgdriveitem -DriveId $TargetOneDrive.Id -DriveItemId "root:/Documents/Cowork/Skills" -erroraction stop
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

Connect-mggraph -Identity

$CoworkSite = get-mgsite -SiteId "$tenantname.sharepoint.com:/sites/$SiteName" 

$Documents = get-mgsitedrive -SiteId $CoworkSite.Id  -Filter "Name eq 'Documents'"

$List = get-mgsiteList -SiteId $CoworkSite.Id -Filter "displayName eq 'Documents'"

[array]$items = Get-MgSiteListItem -SiteId $CoworkSite.Id -ListId $List.id -Select 'fields' -ExpandProperty 'fields'


$Items = Invoke-MgGraphRequest -method Get -Uri "https://graph.microsoft.com/v1.0/sites/$($Coworksite.Id)/lists/$($List.Id)/items?`$expand=fields(`$select=id,Title,DistributeTo,DistributeToLookupId)"

$Deployments = $Items.value

$LicensedUsers = Get-MgUser -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -CountVariable LicensedUsersCount


foreach ($Deployment in $Deployments) {
    $file = Invoke-MgGraphRequest -method Get -Uri "https://graph.microsoft.com/v1.0/sites/$($Coworksite.Id)/lists/$($List.Id)/items/$($deployment.id)/driveitem"
    $sourceFilePath = $file.parentReference.path.Replace("/sites/$SiteName/drive/root:", "") + "/" + $file.name
    
        write-host "Deploying to EEEU Group"
        # Get list of all users in tenant who have licenses
        foreach($licensedUser in $LicensedUsers) {
            Try{
                Get-MgUserDefaultDrive -UserId $licensedUser.UserPrincipalName -ErrorAction Stop | Out-Null
                write-host "Deploying to User: $($licensedUser.UserPrincipalName) via Everyone except external users group"
                CopyFileToOneDrive -TenantName $TenantName -SiteName $SiteName -SourceDriveId $Documents.Id -TargetUserName $licensedUser.UserPrincipalName -SourceFileId $file.id -overwriteExisting:$OverwriteExisting

            }catch{
                write-host "User $($licensedUser.UserPrincipalName) does not have a OneDrive, skipping deployment"
            }
        }
}

