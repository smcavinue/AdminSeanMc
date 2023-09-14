##A report of all email addresses in the Exchange Online tenant
Try {
    get-organizationconfig | out-null
}
catch {
    write-host "You are not connected to Exchange Online, connecting now..."
    Connect-ExchangeOnline
}

Connect-MgGraph -scopes "user.readwrite.all"

$domain = "boatyardx.com"
$onmicrosoftdomain = Get-AcceptedDomain | Where-Object{$_.DomainName -like "*onmicrosoft.com"} | Select-Object -ExpandProperty DomainName
$Recipients = get-EXOrecipient -resultsize unlimited -Properties EmailAddresses, primarysmtpaddress

foreach ($recipient in $recipients) {

    foreach ($address in $recipient.emailaddresses) {
        try{
        $onmicrosoftaddress = ([array]($recipient.emailaddresses | Where-Object{$_ -like "*onmicrosoft.com"}))[0].split(":")[1]
        }catch{}
        if(!$onmicrosoftaddress) {
            $onmicrosoftaddress = $recipient.primarysmtpaddress.Split("@")[0] + "@$onmicrosoftdomain"
        }
        if ($address -like "*$domain*") {  
            switch ($recipient.RecipientType) {
                UserMailbox {
                    $UPN = (Get-EXOMailbox $recipient.Identity).userprincipalname
                    if($upn -eq $address.split(':')[1] ){
                        write-host "changing $upn to $onmicrosoftaddress"
                        $userobject = Get-MgUser -UserId $upn
                        update-mguser -UserId $userobject.Id -userprincipalname $onmicrosoftaddress
                    }
                    set-mailbox -identity $recipient.Identity -WindowsEmailAddress $onmicrosoftaddress
                    Set-Mailbox -Identity $recipient.Identity -EmailAddresses @{remove = $address}
                  }
                  MailUniversalDistributionGroup {
                    Set-DistributionGroup -Identity $recipient.Identity -WindowsEmailAddress $onmicrosoftaddress
                    Set-DistributionGroup -Identity $recipient.Identity -EmailAddresses @{remove = $address}

                }
            }
            
        }

    }

}