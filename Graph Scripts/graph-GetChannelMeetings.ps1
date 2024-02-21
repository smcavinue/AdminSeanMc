connect-mggraph -Scopes Calendars.Read

##Variable to contain all of the team email addresses
$EmailAddresses = @(
    "serviceteam@seanmcavinue.net",
    "PlannerGroup@seanmcavinue.net"
)


##Replace userID with each user you want to check
$Events = get-mguserevent -userid db3a70f1-8e82-4a6d-a936-4695f1f7702a -All
    
    
foreach ($Event in $Events) {
    if ($EmailAddresses -contains $Event.Organizer.EmailAddress.Address) {
        $Event | fl CreatedDateTime, Subject, Organizer
    }else{
        #write-host $event.Organizer.EmailAddress.Address -ForegroundColor red
    }
}




