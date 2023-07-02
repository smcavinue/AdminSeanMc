Connect-MgGraph -Scopes "directory.read.all ChannelSettings.Read.All"

$Teams = Get-MgGroup -Filter "groupTypes/any(c:c eq 'unified') and startswith(displayName,'Migrated-')"

foreach($team in $teams){
    Get-MgTeamChannel -TeamId $team.Id | ForEach-Object {
        $Channel = $_
        $ChannelSettings = Get-MgTeamChannelFileFolder -TeamId $team.Id -ChannelId $Channel.Id
        $ChannelSettings
    }
}
