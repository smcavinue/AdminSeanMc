##Change Graph Version Endpoint
Select-MgProfile v1.0
Select-MgProfile beta

##Verify Graph Version Endpoint
Get-MgProfile

##List Users
Get-MgUser -All

##Create a new user
 $PasswordProfile = @{
   Password = <Password>
   }

New-MgUser -DisplayName <Display Name> -PasswordProfile $PasswordProfile -AccountEnabled -MailNickname <Mail Nickname> -UserPrincipalName <User Principal Name>

##List Groups
Get-MgGroup

##List Group Members
Get-MgGroupMember -GroupId <Group ID>

##Add Group Member
New-MgGroupMember -GroupId <Group ID> -DirectoryObjectId <User ID>

##Remove Group Member
$URI = "https://graph.microsoft.com/v1.0/groups/<Group ID>/members/<User ID>/`$ref"

Invoke-MgGraphRequest `
     -Uri $URI `
     -Method DELETE

##List Teams
Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All

##List Team settings as JSON
Get-MgTeam -TeamId <Group ID> | ConvertTo-Json

##Update Team messaging settings
$messagingSettings = @{
	 "allowUserEditMessages" = "false";
         "allowUserDeleteMessages" = "false" ;

 }

Update-MgTeam -TeamId <Group ID> -MessagingSettings $messagingSettings

##Get Drive
Get-MgGroupDrive -GroupId <Group ID>

##List Drive Items
$URI = "https://graph.microsoft.com/v1.0/Drives/<Drive ID>/root/children

Invoke-MgGraphRequest `
     -Uri $URI `
     -Method GET | ConvertTo-JSON

##List Drive folder items
$URI = "https://graph.microsoft.com/v1.0/Drives/<Drive ID>/items/<Item ID>/children"

(Invoke-MgGraphRequest `
      -Uri $URI `
      -Method GET).value.name

##List All Sites
Get-MgSite -Search "*"

##Search for a specific site
Get-MgSite -Search "<Search Term>"

##Get Site Lists
Get-MgSiteList -SiteId <Site ID>

##Get Site List Columns
Get-MgSiteListColumn -SiteId <Site ID> -ListId <List ID>

##Get mail folder
Get-MgUserMailFolder -UserId 28018250-341c-4a6b-813d-0d76c12c9383

##Get mail folder messages
 Get-MgUserMailFolderMessage -MailFolderId AAMkADMwNDY3OTgyLWI0Y2YtNDg1Ni1iZDgyLTk2NzM0MTQ0MjJlMAAuAAAAAACqvO7TSrJXTIBavsvAPKOXAQCHSvST46vPTrexwpCsCxphAAAAAAEMAAA= -UserId 28018250-341c-4a6b-813d-0d76c12c9383 | fl subject,bodypreview

##Send email message
$MessageDetails = @{
	Message = @{
		Subject = "System automated message"
        importance = "High"
		Body = @{
			ContentType = "html"
			Content = "Test automated message from Graph!"
		}
		ToRecipients = @(
			@{
				EmailAddress = @{
					Address = <To Recipient 1>
				}
            }
            @{
                EmailAddress = @{
                    Address = <To Recipient 2>
                }
			}
		)
		CcRecipients = @(
			@{
				EmailAddress = @{
					Address = <CC Recipient>
				}
			}
		)
	}
	SaveToSentItems = "true"
}

Send-MgUserMail -UserId <Source User>-BodyParameter $MessageDetails

##Get a users calendar
Get-MgUserCalendar -UserId <User ID> | fl

##Get Calendar items
Get-MgUserCalendarEvent -UserId <User ID> -CalendarId <Calendar ID> | fl

##New Calendar Event
$EventDetails= @{
	Subject = "New Appointment from Graph"
	Body = @{
		ContentType = "HTML"
		Content = "This is a new appointment created from Graph"
	}
	Start = @{
		DateTime = "2022-09-13T12:00:00"
		TimeZone = "Pacific Standard Time"
	}
	End = @{
		DateTime = "2022-09-13T14:00:00"
		TimeZone = "Pacific Standard Time"
	}
}

New-MgUserCalendarEvent -UserId <User ID> -CalendarId <Calendar ID> -BodyParameter EventDetails


$ConsentScope = @{
	Scope= ""
}

##Return Service Principal for Microsoft Graph PowerShell
Get-MgServicePrincipal -all | ? {$_.DisplayName -eq "Microsoft Graph PowerShell"}

##Clear consents for Microsoft Graph PowerShell
$ConsentScope = @{
 Scope= ""
 }

Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId 4rrqHK1zwE-NzeI2U278bGMhHt_t0axAkQEYuLbh4MM -BodyParameter  $ConsentScope