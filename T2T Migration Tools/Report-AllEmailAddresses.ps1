##A report of all email addresses in the Exchange Online tenant
Try {
    get-organizationconfig | out-null
}
catch {
    write-host "You are not connected to Exchange Online, connecting now..."
    Connect-ExchangeOnline
}

$Recipients = get-EXOrecipient -resultsize unlimited -Properties EmailAddresses, primarysmtpaddress

foreach ($recipient in $Recipients) {

    foreach ($emailaddress in $recipient.emailaddresses) {
        
    
        $Object = @{
            "Identity"           = $recipient.Identity
            "RecipientType"     = $recipient.RecipientType
            "PrimarySMTPAddress" = $recipient.PrimarySMTPAddress
            "EmailAddresses"     = $emailaddress
        }

        $ExportObject = New-Object -TypeName PSObject -Property $Object

        $ExportObject | Export-Csv -Path "C:\temp\AllEmailAddresses.csv" -Append -NoTypeInformation
    }
}