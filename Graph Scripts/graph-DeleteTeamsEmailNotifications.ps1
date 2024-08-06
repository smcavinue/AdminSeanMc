<#
.SYNOPSIS
This script deletes email notifications for Microsoft Teams from all mailboxes in the tenant.

.DESCRIPTION
This script is used to delete email notifications for Microsoft Teams from Exchange mailboxes. It can be used to remove email notifications from Teams and should be run from an Azure Automation Account Runbook.

.PARAMETER Mailbox
This optional parameter can be used to target a single mailbox

.PARAMETER DeleteOlderThanDays
This optional parameter can be used to delete email notifications older than a specific number of days. The default value is 7 days.

.PARAMETER SearchDays
This optional parameter can be used to search for email notifications in the mailbox for a specific number of days. The default value is 14 days.

.PARAMETER SendReportTo
This optional parameter can be used to send a report to a specific email address. By default, no report is sent. The correct value for this is an email address in your organization.

.PARAMETER ReportOnly
This optional switch can be used to generate a report without deleting any email notifications. The default value is $false.

.PARAMETER ReportFromAddress
This optional parameter is used to specify the from address for the report. If this is not used, the report is not sent.

.EXAMPLE
.\graph-DeleteTeamsEmailNotifications.ps1
Deletes email notifications for all mailboxes. The previous 14 days will be scanned and mails older than 7 days will be deleted.

.EXAMPLE
.\graph-DeleteTeamsEmailNotifications.ps1 -ReportOnly -ReportTo adminseanmc@contoso.com -ReportFromAddress TeamsEmailRemoval@contoso.com
Provides a report to adminseanmc@contoso.com, does NOT delete anything. The previous 14 days will be scanned and mails older than 7 days will be reported upon.

.EXAMPLE
.\graph-DeleteTeamsEmailNotifications.ps1 -Mailbox adminseanmc@contoso.com -SearchDays 30 -DeleteOlderThanDays 14 -SendReportTo adminseanmc@contoso.com  -ReportFromAddress TeamsEmailRemoval@contoso.com
Deletes email notifications for the mailbox adminseanmc@contoso.com. The previous 30 days will be scanned and mails older than 14 days will be deleted. A report mail will be sent to adminseanmc@contoso.com

.NOTES
- This script requires the Microsoft.Graph module to be installed. You can install it by running "Install-Module -Name Microsoft.Graph".
- You need to have the necessary permissions to manage Teams and email notifications .
- You need to have the necessary permissions to manage Teams and email notifications in your organization.
#>

## Parameters
param (
    [Parameter(
        Mandatory = $false,
        HelpMessage = "This optional parameter can be used to target a single mailbox."
    )]
    [string]$Mailbox,

    [Parameter(
        Mandatory = $false,
        HelpMessage = "This optional parameter can be used to delete email notifications older than a specific number of days. The default value is 7 days."
    )]
    [int]$DeleteOlderThanDays = 7,

    [Parameter(
        Mandatory = $false,
        HelpMessage = "This optional parameter can be used to search for email notifications in the mailbox for a specific number of days. The default value is 14 days."
    )]
    [int]$SearchDays = 14,

    [Parameter(
        Mandatory = $false,
        HelpMessage = "This optional parameter can be used to send a report to a specific email address. By default, no report is sent. The correct value for this is an email address in your organization."
    )]
    [string]$SendReportTo,

    [Parameter(
        Mandatory = $false,
        HelpMessage = "This optional switch can be used to generate a report without deleting any email notifications. The default value is `$false."
    )]
    [switch]$ReportOnly,

    [Parameter(
        Mandatory = $false,
        HelpMessage = "This optional parameter is used to specify the from address for the report. If this is not used, the report is not sent.."
    )]
    [string]$ReportFromAddress
    
)

$TeamsEmail = "noreply@emeaemail.teams.microsoft.com"

# Try to Connect to Microsoft Graph
try { 
    Connect-MgGraph -Identity
}
catch {
    write-output "Failed to connect to Microsoft Graph. Please check your credentials and try again."
    return
}

#If no Mailbox is specified, get all users, else get the specified mailbox
If (!$Mailbox) {
    
    $Users = Get-MGUser -All
}
Else {
    [array]$Users = Get-MgUser -UserId $Mailbox
}
# If the user has a mailbox, add to array for processing
$Mailboxes = @()
foreach ($user in $users) {
    Try {
        $MailboxSettings = Get-MgUser -UserId $user.Id -Select mailboxSettings -ErrorAction Stop
        write-output "Mailbox found for $($user.UserPrincipalName)"
        $Mailboxes += $user

    }
    Catch {
        write-output "No Mailbox found for $($user.UserPrincipalName)"
        Continue
    }
}

# If no mailboxes are found, exit
If ($Mailboxes.Count -eq 0) {
    write-output "No mailboxes found. Exiting."
    return
}

write-output "Processing $($Mailboxes.Count) mailboxes"

# Prepare Table
$EmailTable = "<table>"

# Loop through each mailbox
foreach ($MailboxObject in $Mailboxes) {
    write-output "Starting to process mailbox $($MailboxObject.UserPrincipalName)"

    # Get emails from the defined timeframe
    $StartDate = (Get-Date).AddDays(-$SearchDays).ToUniversalTime()
    $StartDate = get-date $startdate -UFormat '+%Y-%m-%dT%H:%M:%S.000Z'

    $EndDate = (Get-Date).ToUniversalTime()
    $EndDate = get-date $enddate -UFormat '+%Y-%m-%dT%H:%M:%S.000Z'

    [array]$Emails = (Get-MgUserMessage -UserId  $mailboxObject.id -Filter "ReceivedDateTime ge $($StartDate.tostring()) and ReceivedDateTime le $EndDate and sender/emailAddress/address eq '$TeamsEmail'" -All)

    # If no emails are found, exit
    If ($Emails.Count -eq 0) {
        write-output "No emails found for $($MailboxObject.UserPrincipalName). Exiting."
        Continue
    }

    # Delete mails (if applicable) and generate a HTML Table of all emails found containing mailbox, subject, received date and deletion status
    $EmailTable += "<tr><th>Mailbox</th><th>Subject</th><th>Received Date</th><th>Status</th></tr>"
    foreach ($Email in $Emails) {
        if ($ReportOnly) {
            $status = "Report Only"
        }else{
            Try{
                Remove-MgUserMessage -UserId $MailboxObject.id -MessageId $Email.Id -ErrorAction Stop
                $status = "Deleted"
            }
            Catch {
                $status = "Failed to Delete"
            
            }
        }
        $EmailTable += "<tr><td>$($MailboxObject.UserPrincipalName)</td><td>$($Email.Subject)</td><td>$($Email.ReceivedDateTime)</td><td>$($Status)</td></tr>"
    }
}
$EmailTable += "</table>"

#Apply formatting and borders to the EmailTable
$EmailTable = $EmailTable -replace "<table>", "<table border='1' style='border-collapse: collapse;'>"

# Send the report by email
if ($SendReportTo) {
    $ReportSubject = "Teams Email Notifications Report"
    $ReportBody += "<br><br>Report generated on $(Get-Date)"
    $ReportBody += $EmailTable
    $ReportBody += "<br><br>This report was generated by the Teams Email Notifications Removal Script."
    $ReportBody += "<br><br>Thank you."

    if ($ReportOnly) {
        $ReportBody += "<br><br>Report Only Mode Enabled - No deletion has occurred."
    }
    else {
        $ReportBody += "<br><br>Report Only Mode Disabled - Mails have been Deleted."    
    }

    $params = @{
        Message         = @{
            Subject       = $Reportsubject
            Body          = @{
                ContentType = "HTML"
                Content     = $Reportbody
            }
            ToRecipients  = @(
                @{
                    EmailAddress = @{
                        Address = $SendReportTo
                    }
                }
            )
        }
    }
    ##Send the report in email
    Try{
        
        Send-MgUserMail  -UserId $ReportFromAddress -BodyParameter $params
        write-output  "Mail sent successfully to $SendReportTo"
    }catch{
        write-output  "Error sending mail to $SendReportTo"
    }
        
}

