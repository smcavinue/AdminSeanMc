Param
(

  [Parameter (Mandatory= $true)]
  [String] $SiteURL,
  
  [Parameter (Mandatory= $true)]
  [String] $adminURL,

  [Parameter (Mandatory= $true)]
  [String] $GroupName,

  [Parameter (Mandatory= $true)]
  [String] $GroupDescription,

  [Parameter (Mandatory= $true)]
  [String] $DefaultOwner1,
  
  [Parameter (Mandatory= $true)]
  [String] $DefaultOwner2,

  [Parameter (Mandatory= $true)]
  [bool] $InternalOnly
)

##Site URL will be
$URL = "$($adminURL.replace('-admin',''))/sites/$siteURL"


##Connect to Azure to retrieve an access token
Connect-AzAccount -Identity
$Token = (Get-AzAccessToken -ResourceURL "https://graph.microsoft.com").token


##Connect to Microsoft Graph
Connect-MgGraph -AccessToken $Token
select-mgprofile beta

##Build Parameters for new M365 Group
$GroupParam = @{
    DisplayName = $GroupName
    description = $GroupDescription
    GroupTypes = @(
        "Unified"
    )
    SecurityEnabled     = $false
    IsAssignableToRole  = $false
    MailEnabled         = $false
    MailNickname        = $SiteURL
    "Owners@odata.bind" = @(
        "https://graph.microsoft.com/v1.0/users/$DefaultOwner1",
        "https://graph.microsoft.com/v1.0/users/$DefaultOwner2"
    )
}

#Provision M365 Group
$Group = New-MgGroup -BodyParameter $GroupParam


##Wait for Group to finish provisioning
start-sleep -seconds 60

##Build Parameters for new Team

$TeamParam = @{
	"Template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
	"Group@odata.bind" = "https://graph.microsoft.com/v1.0/groups('$($Group.id)')"
	Channels = @(
		@{
			DisplayName = "Welcome!"
			IsFavoriteByDefault = $true
		}
	)
	MemberSettings = @{
		AllowCreateUpdateChannels = $false
		AllowDeleteChannels = $false
		AllowAddRemoveApps = $false
		AllowCreateUpdateRemoveTabs = $false
		AllowCreateUpdateRemoveConnectors = $false
	}
}

##Add Team to group
New-MgTeam -BodyParameter $TeamParam

##Add Group to tenant group lifecycle policy
Add-MgGroupToLifecyclePolicy -GroupLifecyclePolicyId (Get-MgGroupLifecyclePolicy).id -GroupId $Group.id

##If the Team is internal only, block guest access and external sharing
if($InternalOnly){

##Block Guest Access to Group if required
$Template = Get-MgDirectorySettingTemplate | ?{$_.displayname -eq "Group.Unified.Guest"}

$TemplateParams = @{
	TemplateId = "$($template.id)"
	Values = @(
		@{
			Name = "AllowToAddGuests"
			Value = "false"
		}
	)
}
New-MgGroupSetting -BodyParameter $TemplateParams -GroupId $Group.id


##Connect to PNP
Connect-PnPOnline -ManagedIdentity -Url $adminURL

##Set Sharing policy to internal only
$site = Set-pnptenantsite -SharingCapability Disabled -Url $url


}