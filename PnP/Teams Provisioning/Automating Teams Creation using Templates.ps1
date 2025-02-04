##Team Settings
$DisplayName = "Project - Project Spring"
$Description = "Project Spring is a project team for the Spring project"
$MailNickName = "Project-ProjectSpring" ##Mail Nickname is also used as site URL
$ClientID = "<ClientID>"
$tenantName = "<TenantName>"

#Connect PNP Online
Connect-pnpOnline $tenantName -Interactive -ClientId $ClientID

##Project Team Settings
$ProjectTeamSettings = @{
    AllowDeleteChannels = $false
    AllowAddRemoveApps = $true
    AllowChannelMentions = $true
    AllowCreateUpdateConnectors = $true
    AllowCustomMemes = $true
    AllowGiphy = $true
    AllowStickersAndMemes = $true
    AllowTeamMentions = $true
    GiphyContentRating = "Moderate"
    AllowUserEditMessages = $true
    AllowOwnerDeleteMessages = $false
    AllowCreateUpdateChannels = $true
    AllowCreateUpdateRemoveTabs = $true
    AllowUserdeleteMessages = $false
    AllowGuestCreateUpdateChannels = $false
    AllowGuestDeleteChannels = $false
    visibility = "Private"
    SensitivityLabel = "258c86d0-ca73-4202-a8d5-d01dd9abaf80"
}

$TeamObject = New-PnPTeamsTeam -DisplayName $DisplayName -Description $Description -MailNickName $MailNickName -AllowDeleteChannels $ProjectTeamSettings.AllowDeleteChannels -AllowAddRemoveApps $ProjectTeamSettings.AllowAddRemoveApps -AllowCreateUpdateChannels $ProjectTeamSettings.AllowCreateUpdateChannels -AllowCreateUpdateRemoveTabs $ProjectTeamSettings.AllowCreateUpdateRemoveTabs -AllowUserdeleteMessages $ProjectTeamSettings.AllowUserdeleteMessages -AllowUsereditMessages $ProjectTeamSettings.AllowUsereditMessages -AllowGuestCreateUpdateChannels $ProjectTeamSettings.AllowGuestCreateUpdateChannels -AllowGuestDeleteChannels $ProjectTeamSettings.AllowGuestDeleteChannels -Visibility $ProjectTeamSettings.visibility -SensitivityLabel $ProjectTeamSettings.SensitivityLabel

##Team structures
##Project Team structure
$ProjectTeamChannels = @(
    @{
        DisplayName = "Project Discussion"
        Description = "General discussion channel for team"
        ChannelType = "Standard"
    },
    @{
        DisplayName = "Project Plan"
        Description = "Channel for project plan"
        ChannelType = "Standard"
    },
    @{
        DisplayName = "Project Scheduling"
        Description = "Channel for discussing project scheduling"
        ChannelType = "Standard"
    },
    @{
        DisplayName = "Execution"
        Description = "Channel for execution of the project"
        ChannelType = "Standard"
    },
    @{
        DisplayName = "Project Status Reports"
        Description = "Channel for project status reports"
        ChannelType = "Standard"
    },
    @{
        DisplayName = "Project Budget Tracking"
        Description = "Private channel for project budget and financials"
        ChannelType = "Private"
    }
)

$Owner = (Get-PnPTeamsUser -Team $TeamObject.GroupId).userprincipalname
foreach($Channel in $ProjectTeamChannels)
{
    if($Channel.ChannelType -eq "Private")
    {
        Add-PnPTeamsChannel -Team $TeamObject.GroupId -DisplayName $Channel.DisplayName -Description $Channel.Description -ChannelType $Channel.ChannelType -OwnerUPN $Owner
    }
    else
    {
        Add-PnPTeamsChannel -Team $TeamObject.GroupId -DisplayName $Channel.DisplayName -Description $Channel.Description -ChannelType $Channel.ChannelType
    }
}

$SiteURL = Get-PnPMicrosoft365Group -Identity $TeamObject.GroupId -IncludeSiteUrl | Select-Object -ExpandProperty SiteUrl

connect-PnPOnline -Interactive -ClientId $ClientID -Url $SiteURL
Invoke-PnPSiteTemplate -Path C:\SiteTemplates\TeamSiteTemplate.pnp 