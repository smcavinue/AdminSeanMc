##Get Site Role Assignments
#Get all the role assignments for the site
$RoleAssignments = (Get-PnPWeb -Includes RoleAssignments).RoleAssignments

#Loop through each role assignment
foreach($RoleAssignment in $RoleAssignments)
{
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
foreach($GroupMember in $GroupMembers)
{
    if($GroupMember.PrincipalType -eq "SecurityGroup")
    {
        $NestedGroupMembers = Get-PnPEntraIDGroupMember -Identity $GroupMember.LoginName.Split('|')[-1] 
        foreach($NestedGroupMember in $NestedGroupMembers)
        {
            Write-Host "Nested Group Member: $($NestedGroupMember.displayName) is a member of $($GroupMember.Title)"
        }
    }
        Write-Host "Owner Group Member: $($GroupMember.Title)"
}
