Connect-ExchangeOnline
Connect-MsolService
<#
    Author: Sean McAvinue
    Contact: https://seanmcavinue.net, Twitter: @Sean_McAvinue
    .SYNOPSIS
    Export CSV report of all mailbox sizes and assigned licenses


#>
$mailboxes = Get-EXOMailbox -ResultSize unlimited

foreach($mailbox in $mailboxes){
 
    $stats = Get-EXOMailboxStatistics $mailbox.identity -ErrorAction SilentlyContinue
    $MSOLAccount = get-msoluser -UserPrincipalName $mailbox.UserPrincipalName -ErrorAction SilentlyContinue

    $ExportObject = [PSCustomObject]@{
        PrimarySMTPAddress = $mailbox.primarysmtpaddress
        UserPrincipalName = $MSOLAccount.userprincipalname
        Mailboxsize = $stats.Totalitemsize
        LicensesAssigned = $MSOLAccount.Licenses.accountskuid -join ';'
    }

    $ExportObject | export-csv MailboxReport.csv -NoClobber -NoTypeInformation -Append
}