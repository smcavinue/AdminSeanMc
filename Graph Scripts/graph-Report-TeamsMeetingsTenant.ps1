$ClientID = "<Your Client ID Here"
$TenantID = "<Your Tenant ID Here"
$CertificateThumbprint ="<Your Certificate Thumbprint Here>"

Connect-MgGraph -ClientID $ClientID -TenantId $TenantID -CertificateThumbprint $CertificateThumbprint

##Show User list
$User = get-mguser -All | Out-GridView -PassThru

##Get all future calendar items in the users mailbox
[array]$CalendarItems = Get-MgUserEvent -UserId $User.Id -All | ?{[datetime]$_.Start.DateTime -gt (Get-Date) -or $_.Recurrence.Range.EndDate -gt (Get-Date) -or $_.Recurrence.Range.Type -eq "NoEnd"}

##Get the Teams Meeting Items
[array]$TeamsMeetings = $calendarItems | ?{$_.Body.Content -like "*https://teams.microsoft.com/l/meetup-join*"}

##Extract the meeting options URL from the body content
foreach($meeting in $TeamsMeetings){
    ##Remove everything before the URL
    $MeetingTenantID = (($meeting.Body.Content.tostring() -split 'tenantId=')[-1] -split "&amp;threadId=")[0]
    $TenantDetails = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/tenantRelationships/findTenantInformationByTenantId(tenantId='$MeetingTenantID')")
    $Meeting | Add-Member -MemberType NoteProperty -Name TenantDetails -Value $TenantDetails.tenantId -Force
    $Meeting | Add-Member -MemberType NoteProperty -Name TenantDomainName -Value $TenantDetails.defaultDomainName -Force
    $Meeting | Add-Member -MemberType NoteProperty -Name TenantDisplayName -Value $TenantDetails.displayName -Force
    $meeting | export-csv c:\temp\$($User.UserPrincipalName)_TeamsMeetings.csv -NoTypeInformation -Append

}
