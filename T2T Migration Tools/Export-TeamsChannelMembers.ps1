$BYXTeams = import-csv C:\temp\Teams.csv

foreach($team in $BYXTeams){

$group = Get-UnifiedGroup $team.'Source MailNickname'
#write-host $group.ExternalDirectoryObjectId.Split(' ')[0]
$Teamobject = Get-Team -GroupId $group.ExternalDirectoryObjectId.Split(' ')[0]
[array]$Channels = Get-TeamChannel -GroupId $Teamobject.GroupId | ?{$_.membershiptype -ne "standard"}

foreach($channel in $channels){

    [array]$Members = Get-TeamChannelUser -GroupId $Teamobject.GroupId -DisplayName $channel.displayname -Role owner | ?{$_.user -like "*@boatyardxsolutions.onmicrosoft.com"}

    foreach($member in $members){

        write-host "$($member.user) is a member of $($channel.displayname) in Team $($teamobject.displayname)"

          $Object = @{
            "UPN"           = $member.User
            "Group"     = $team.NewName
            "channel" = $channel.DisplayName
            "role"     = "owner"
        }

        $ExportObject = New-Object -TypeName PSObject -Property $Object

        $ExportObject | export-csv C:\temp\Teamschannelmembers.csv -NoClobber -NoTypeInformation -Append

    }

}

}