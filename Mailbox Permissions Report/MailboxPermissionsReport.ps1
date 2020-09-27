#Retrieve all mailboxes
$mailboxes = get-Exomailbox -ResultSize unlimited

#Loop through mailboxes
foreach($mailbox in $mailboxes){
    
    #Get Mailbox Permissions
    $Permissions = Get-EXOMailboxPermission $mailbox.primarysmtpaddress

    #Loop Through Permission set
    foreach($permission in $permissions){


    #Build Object
    $object = [pscustomobject]@{
                'UserEmail' = $permission.user
                'MailboxEmail' = $mailbox.primarysmtpaddress
                'PermissionSet' = [String]$permission.accessrights
                }
    
    #Export Details
    $Object | export-csv MailboxPermissions.csv -NoClobber -NoTypeInformation -Append 

    #Remove Object
    Remove-Variable object

    }



}