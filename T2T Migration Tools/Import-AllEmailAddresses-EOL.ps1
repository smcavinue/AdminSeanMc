##A report of all email addresses in the Exchange Online tenant
Try {
    get-organizationconfig | out-null
}
catch {
    write-host "You are not connected to Exchange Online, connecting now..."
    Connect-ExchangeOnline
}


$csv = import-csv C:\temp\AllEmailAddresses.csv

foreach ($mailbox in $csv) {

    $addresstoadd = $mailbox.EmailAddresses.split(":")[1]

    if ($mailbox.emailaddresses -clike "SMTP:*") {
        $isprimary = $true
    }
    else {
        $isprimary = $false
    }

    switch ($mailbox.RecipientType) {
        UserMailbox { 
            if ($isprimary) {
                write-host "Adding PRIMARY ADDRESS to MAILBOX $addresstoadd to $($mailbox.targetprimarysmtpaddress)"
                set-mailbox -identity $mailbox.targetprimarysmtpaddress -WindowsEmailAddress $addresstoadd
            }
            else {
                write-host "Adding ALIAS to MAILBOX $addresstoadd to $($mailbox.targetprimarysmtpaddress)"
                set-mailbox -identity $mailbox.targetprimarysmtpaddress -EmailAddresses @{add = $addresstoadd }
            }
        }
        MailUniversalDistributionGroup {
            if ($isprimary) {

                write-host "Adding PRIMARY ADDRESS to GROUP $addresstoadd to $($mailbox.targetprimarysmtpaddress)"
                set-distributiongroup -identity $mailbox.targetprimarysmtpaddress -WindowsEmailAddress $addresstoadd
            }
            else {
                
                write-host "Adding ALIAS to GROUP $addresstoadd to $($mailbox.targetprimarysmtpaddress)"
                set-distributiongroup  -identity $mailbox.targetprimarysmtpaddress -EmailAddresses @{add = $addresstoadd }
            }
        }
    }
}