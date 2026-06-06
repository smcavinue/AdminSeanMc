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
    $TargetFilePath = "root:/Documents/Cowork2/Skills"
    $TargetSkill = "root:/Documents/Cowork2/Skills/$($SourceFile.Name)"


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

$Deployments = $Items.value | Where-Object { $_.fields.DistributeTo -Ne $null }

$PeopleList = get-mgsiteList -SiteId $CoworkSite.Id -Filter "displayName eq 'User Information List'"

foreach ($Deployment in $Deployments) {
    $file = Invoke-MgGraphRequest -method Get -Uri "https://graph.microsoft.com/v1.0/sites/$($Coworksite.Id)/lists/$($List.Id)/items/$($deployment.id)/driveitem"
    $sourceFilePath = $file.parentReference.path.Replace("/sites/$SiteName/drive/root:", "") + "/" + $file.name
    $User = Invoke-MgGraphRequest -method Get -Uri "https://graph.microsoft.com/v1.0/sites/$($Coworksite.Id)/lists/$($PeopleList.Id)/items/$($Deployment.fields.DistributeToLookupId)?expand=fields"
    
    If ($Deployment.fields.DistributeTo -eq "Everyone except external users") {
        write-host "Deployed to EEEU Group"
        # Get list of all users in tenant who have licenses
        $LicensedUsers = Get-MgUser -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -CountVariable LicensedUsersCount
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
    elseif ($User.fields.ContentType -eq "Person") {
        write-host "Deployed to User: $($User.fields.username)"
        Try{
            Get-MgUserDefaultDrive -UserId $User.fields.username -ErrorAction Stop | Out-Null
            CopyFileToOneDrive -TenantName $TenantName -SiteName $SiteName -SourceDriveId $Documents.Id -TargetUserName $User.fields.username -SourceFileId $file.id -overwriteExisting:$OverwriteExisting
    }catch{
            write-host "User $($User.fields.username) does not have a OneDrive, skipping deployment"
        }
    }
    elseIf ($User.fields.ContentType -eq "DomainGroup") {
        write-host "Deployed to Group: $($Deployment.fields.DistributeTo)"
        [array]$GroupMembers = (Get-MgGroupMember -GroupId $user.fields.name.split('|')[-1]).additionalproperties.userPrincipalName

        foreach($GroupMember in $GroupMembers) {
            Try{
                Get-MgUserDefaultDrive -UserId $GroupMember -ErrorAction Stop | Out-Null
                write-host "Deploying to User: $GroupMember via Group: $($Deployment.fields.DistributeTo)"
                CopyFileToOneDrive -TenantName $TenantName -SiteName $SiteName -SourceDriveId $Documents.Id -TargetUserName $GroupMember -SourceFileId $file.id -overwriteExisting:$OverwriteExisting

            }catch{
                write-host "User $GroupMember does not have a OneDrive, skipping deployment"
            }
        }

    }
}

