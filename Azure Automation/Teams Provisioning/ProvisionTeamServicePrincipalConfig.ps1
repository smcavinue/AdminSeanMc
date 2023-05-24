$GraphApp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0000-c000-000000000000'" # Microsoft Graph
[array]$Roles = $GraphApp.AppRoles | Where-Object {$Permissions -contains $_.Value}
##Loop through and add each role
foreach($role in $roles){
    $AppRoleAssignment = @{
        "PrincipalId" = $MIID
        "ResourceId" = $GraphApp.Id
        "AppRoleId" = $Role.Id 
    }
    # Assign the Graph permission
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $MIID -BodyParameter $AppRoleAssignment
}
##List Required Permissions
$Permissions = @(
    "Sites.FullControl.All"
)
##Get Roles for Permissions
$GraphApp = Get-MgServicePrincipal -Filter "AppId eq '00000003-0000-0ff1-ce00-000000000000'" # SharePoint Online
[array]$Roles = $GraphApp.AppRoles | Where-Object {$Permissions -contains $_.Value}
##Loop through and add each role
foreach($role in $roles){
    $AppRoleAssignment = @{
        "PrincipalId" = $MIID
        "ResourceId" = $GraphApp.Id
        "AppRoleId" = $Role.Id 
    }
    # Assign the SharePoint permission
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $MIID -BodyParameter $AppRoleAssignment
}
