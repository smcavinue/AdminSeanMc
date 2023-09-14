$csv = import-csv C:\temp\Teamschannelmembers.csv


foreach($member in $csv){

    $Group = Get-UnifiedGroup $member.group
    $team = Get-Team -GroupId $group.ExternalDirectoryObjectId
    write-host "checking for $($member.group) channel $($member.channel)"
    $mailbox = Get-Mailbox $member.UPN
    $mailbox.primarysmtpaddress
    Add-TeamChannelUser -GroupId $team.GroupId -DisplayName $member.channel -User $mailbox.ExternalDirectoryObjectId

}

$owners = $csv | ?{$_.role -eq "owner"}

foreach($member in $owners){

    $Group = Get-UnifiedGroup $member.group
    $team = Get-Team -GroupId $group.ExternalDirectoryObjectId
    write-host "checking for OWNER $($member.group) channel $($member.channel)"
    $mailbox = Get-Mailbox $member.UPN
    $mailbox.primarysmtpaddress
    Add-TeamChannelUser -GroupId $team.GroupId -DisplayName $member.channel -User $mailbox.ExternalDirectoryObjectId -Role owner

}
