Connect-MgGraph -Scopes "Group.Read.All ChannelMember.Read.All"

$Date = (get-date).tostring('yyyy-MM-dd_HH-mm-ss')
$Teams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All
$i = 0
Foreach ($Team in $Teams) {
    $i++
    write-progress -activity "Processing Teams" -status "Processing Team $i of $($Teams.Count)" -percentcomplete (($i / $Teams.Count) * 100)


    [array]$TeamOwners = Get-MgGroupOwner -GroupId $Team.Id
    [array]$TeamMembers = Get-MgGroupMember -GroupId $Team.Id
    $Department = (Get-MgUser -UserId $TeamOwners[0].id -Select department).department
    $TeamChannels = (Get-MgTeamChannel -TeamId $Team.Id -Filter "membershipType eq 'standard'").DisplayName -join ';'
    [array]$TeamPrivateChannels = Get-MgTeamChannel -TeamId $Team.Id -Filter "membershipType eq 'private'"
    [array]$TeamSharedChannels = Get-MgTeamChannel -TeamId $Team.Id -Filter "membershipType eq 'shared'"

    $TeamPrivateChannelOutput = @()
    $TeamSharedChannelOutput = @()

    foreach($channel in $TeamPrivateChannels){
        
        [array]$Owners = Get-MgTeamChannelMember -TeamId $Team.Id -ChannelId $channel.Id | ?{$_.roles -like "*Owner*"}
        $OwnersOutput = $Channel.DisplayName + "- Owners: (" + ($Owners.DisplayName -join ';') + ") "
        $TeamPrivateChannelOutput += $OwnersOutput
    }

    foreach($channel in $TeamSharedChannels){
        
        [array]$Owners = Get-MgTeamChannelMember -TeamId $Team.Id -ChannelId $channel.Id | ?{$_.roles -like "*Owner*"}
        $OwnersOutput = $Channel.DisplayName + "- Owners: (" + ($Owners.DisplayName -join ';') + ") "
        $TeamSharedChannelOutput += $OwnersOutput
    }
    
    ##Create a new Object to store results
    $OutputObject = [PSCustomObject]@{
        TeamName = $Team.DisplayName
        TeamDepartment = $Department
        TeamOwners = $TeamOwners.AdditionalProperties.mail -join ';'
        TeamMembers = $TeamMembers.AdditionalProperties.mail -join ';'
        TeamChannels = $TeamChannels
        TeamPrivateChannels = $TeamPrivateChannelOutput -join ';'
        TeamSharedChannels = $TeamSharedChannelOutput -join ';'
    }

    ##Export Object to CSV

    $OutputObject | export-csv -Path "C:\temp\$($Date)TeamsDetailsReport.csv" -NoTypeInformation -Append
}