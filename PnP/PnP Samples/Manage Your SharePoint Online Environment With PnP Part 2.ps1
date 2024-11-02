##Get Site Role Assignments
#Get all the role assignments for the site
$RoleAssignments = (Get-PnPWeb -Includes RoleAssignments).RoleAssignments

#Loop through each role assignment
foreach ($RoleAssignment in $RoleAssignments) {
    #Get the role definition bindings
    $RoleDefinitionBindings = Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings
    #Get the member details
    $Member = Get-PnPProperty -ClientObject $RoleAssignment -Property member
    #Output the role assignment and role definition
    Write-Host "$($member.GetType().name): $($Member.Title) - Role: $($RoleDefinitionBindings.Name)"
}
 


##Get Owner Group
Get-PnPGroup -AssociatedOwnerGroup
##Get Member Group
Get-PnPGroup -AssociatedMemberGroup
##Get Visitor Group
Get-PnPGroup -AssociatedVisitorGroup

##Get Members of the Owner Group
Get-PnPGroup -AssociatedOwnerGroup | Get-PnPGroupMember

##Get Members of the Member Group
Get-PnPGroup -AssociatedMemberGroup | Get-PnPGroupMember

##Get Members of the Visitor Group
Get-PnPGroup -AssociatedVisitorGroup | Get-PnPGroupMember


##Get Member group members and expand any nested groups
$Group = Get-PnPGroup -AssociatedMemberGroup
$GroupMembers = Get-PnPGroupMember -Identity $Group.Id
foreach ($GroupMember in $GroupMembers) {
    if ($GroupMember.PrincipalType -eq "SecurityGroup") {
        $NestedGroupMembers = Get-PnPEntraIDGroupMember -Identity $GroupMember.LoginName.Split('|')[-1] 
        foreach ($NestedGroupMember in $NestedGroupMembers) {
            Write-Host "Nested Group Member: $($NestedGroupMember.displayName) is a member of $($GroupMember.Title)"
        }
    }
    Write-Host "Owner Group Member: $($GroupMember.Title)"
}


##Add current user as a site collection administrator
$currentUser = (Get-PnPProperty -ClientObject (Get-PnPWeb) -Property CurrentUser).LoginName
Set-PnPTenantSite -Url "https://<TenantDomain>.sharepoint.com/sites/demo-finance" -Owners $currentUser

##Add a user to a built in group
$VisitorsGroup = Get-PnPGroup -AssociatedVisitorGroup
Add-PnPGroupMember -LoginName "dale.cooper@contoso.com" -Group $VisitorsGroup.LoginName

##Add new group to a site
New-PnPSiteGroup -Name "Contributors Group" -PermissionLevels "Contribute"

##Create custom permission level and add a user to it
Add-PnPRoleDefinition -RoleName "List Updater" -Clone "Read" -Include EditListItems
Set-PnPWebPermission -User "laura.palmer@contoso.com" -AddRole "List Updater"

##Remove specific Group Member from visitors group
Remove-PnPGroupMember -Group (Get-PnPGroup -AssociatedVisitorGroup).id -LoginName dale.cooper@contoso.com

##Remove all members from the built in memnbers group
$MembersGroup = Get-PnPGroup -AssociatedMemberGroup
$Members = Get-PnPGroupMember -Identity $MembersGroup.Id
foreach ($Member in $Members) {
    Remove-PnPGroupMember -Group $MembersGroup.Title -LoginName $Member.LoginName
}


##Using Connections
$AdministrationConnection = Connect-PnPOnline -Interactive -ClientId "<ClientID>" -Url "https://<TenantDomain>.sharepoint.com/sites/demo-Administration" -ReturnConnection
$FinanceConnection = Connect-PnPOnline -Interactive -ClientId "<ClientID>" -Url "https://<TenantDomain>.sharepoint.com/sites/demo-Finance" -ReturnConnection
Get-PnPWeb -Connection $AdministrationConnection
Get-PnPWeb -Connection $FinanceConnection


##Create Connections
$AdministrationConnection = Connect-PnPOnline -Interactive -ClientId "<ClientID>" -Url "https://<TenantDomain>.sharepoint.com/sites/demo-Administration" -ReturnConnection
$FinanceConnection = Connect-PnPOnline -Interactive -ClientId "<ClientID>" -Url "https://<TenantDomain>.sharepoint.com/sites/demo-Finance" -ReturnConnection
##Get Members from the source site
$SourceMemberGroup = Get-PnPGroup -AssociatedMemberGroup -Connection $AdministrationConnection
$SourceMembers = Get-PnPGroupMember -Group $SourceMemberGroup -Connection $AdministrationConnection
##Get the target site member group
$TargetMemberGroup = Get-PnPGroup -AssociatedMemberGroup -Connection $FinanceConnection
##Add the members to the target site
foreach ($SourceMember in $SourceMembers) {
    Add-PnPGroupMember -LoginName $SourceMember.LoginName -Group $TargetMemberGroup.LoginName -Connection $FinanceConnection
}