Connect-ExchangeOnline

$Recipients = Get-ExoRecipient -ResultSize Unlimited

foreach($recipient in $recipients){

    foreach($address in [array]$recipient.emailaddresses){

                ##new hashtable
                $hash = [ordered]@{
                    "PrimaryAddress" = $recipient.primarysmtpaddress
                    "Address" = $address
                    "RecipientType" = $recipient.recipienttype
                }

                
                ##Convert hashtable to PSCustomObject
                $obj = New-Object -TypeName psobject -Property $hash

                ##Export to CSV
                $obj | Export-Csv -Path "C:\Temp\AllEmailAddresses.csv" -Append -NoTypeInformation

    }

}
