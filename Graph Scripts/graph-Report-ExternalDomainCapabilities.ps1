##Connect to MG Graph and Teams
Connect-MgGraph -Scopes "IdentityProvider.Read.All policy.read.all CrossTenantInformation.ReadBasic.All SharePointTenantSettings.Read.All"
Select-MgProfile beta
Connect-MicrosoftTeams

##Create output object array
$domainSettingsObjectArray = @()

##Get external domains from B2BManagementPolicy
$B2BManagementPolicy = ((Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/legacy/policies" ).value.definition | convertfrom-json).b2bmanagementpolicy
$allowedGuestDomains = $B2BManagementPolicy.InvitationsAllowedAndBlockedDomainsPolicy.alloweddomains
$blockedGuestDomains = $B2BManagementPolicy.InvitationsAllowedAndBlockedDomainsPolicy.blockeddomains

if ((!$allowedGuestDomains) -and (!$blockedGuestDomains)) {
    ##If there are no allow or block lists, default behavior is to allow all domains
    $domainSettingsObject = [pscustomobject]@{
        tenantID                     = ""
        Domain                       = "Default"
        GuestInvitations             = "Allowed"
        TeamsFederation              = ""
        SharePointSharing            = ""
        B2BCrossTenantAccessInbound  = ""
        B2BCrossTenantAccessOutbound = ""
        B2BDirectConnectInbound      = ""
        B2BDirectConnectOutbound     = ""
        CrossTenantSyncEnabled       = ""
        TrustSettings                = ""
        AutomaticRedemption          = ""
    }
    $domainSettingsObjectArray += $domainSettingsObject

}elseif ((!$allowedGuestDomains) -and ($blockedGuestDomains)) {
    ##If there is a block list, default behavior is to allow all domains except those in the block list
    $domainSettingsObject = [pscustomobject]@{
        tenantID                     = ""
        Domain                       = "Default"
        GuestInvitations             = "Allowed"
        TeamsFederation              = ""
        SharePointSharing            = ""
        B2BCrossTenantAccessInbound  = ""
        B2BCrossTenantAccessOutbound = ""
        B2BDirectConnectInbound      = ""
        B2BDirectConnectOutbound     = ""
        CrossTenantSyncEnabled       = ""
        TrustSettings                = ""
        AutomaticRedemption          = ""
    }
    $domainSettingsObjectArray += $domainSettingsObject

    ##Loop through Blocked Domains and add to output object array
    foreach ($domain in $blockedGuestDomains) {
        $domainSettingsObject = [pscustomobject]@{
            tenantID                     = ""
            Domain                       = $domain
            GuestInvitations             = "Blocked"
            TeamsFederation              = ""
            SharePointSharing            = ""
            B2BCrossTenantAccessInbound  = ""
            B2BCrossTenantAccessOutbound = ""
            B2BDirectConnectInbound      = ""
            B2BDirectConnectOutbound     = ""
            CrossTenantSyncEnabled       = ""
            TrustSettings                = ""
            AutomaticRedemption          = ""
        }
        $domainSettingsObjectArray += $domainSettingsObject
    }
}elseif (($allowedGuestDomains) -and (!$blockedGuestDomains)) {	
    ##If there is an allow list, default behavior is to block all domains except those in the allow list
    $domainSettingsObject = [pscustomobject]@{
        tenantID                     = ""
        Domain                       = "Default"
        GuestInvitations             = "Blocked"
        TeamsFederation              = ""
        SharePointSharing            = ""
        B2BCrossTenantAccessInbound  = ""
        B2BCrossTenantAccessOutbound = ""
        B2BDirectConnectInbound      = ""
        B2BDirectConnectOutbound     = ""
        CrossTenantSyncEnabled       = ""
        TrustSettings                = ""
        AutomaticRedemption          = ""
    }
    $domainSettingsObjectArray += $domainSettingsObject
    ##Loop through Allowed Domains and add to output object array
    foreach ($domain in $allowedGuestDomains) {
        $domainSettingsObject = [pscustomobject]@{
            tenantID                     = ""
            Domain                       = $domain
            GuestInvitations             = "Allowed"
            TeamsFederation              = ""
            SharePointSharing            = ""
            B2BCrossTenantAccessInbound  = ""
            B2BCrossTenantAccessOutbound = ""
            B2BDirectConnectInbound      = ""
            B2BDirectConnectOutbound     = ""
            CrossTenantSyncEnabled       = ""
            TrustSettings                = ""
            AutomaticRedemption          = ""
        }
        $domainSettingsObjectArray += $domainSettingsObject
    }
}

##Get SharePoint Online Tenant Settings
$Uri = "https://graph.microsoft.com/beta/admin/sharepoint/settings"
$SPOSettings = Invoke-MgGraphRequest -Uri $Uri -Method Get
if($SPOSettings.sharingCapability -eq "Disabled"){
    $domainSettingsObjectArray | ? { $_.domain -eq "Default" } | % { $_.SharePointSharing = "Blocked" }
}elseif(($SPOSettings.sharingCapability -ne "Disabled") -and ($SPOSettings.sharingDomainRestrictionMode -eq "AllowList")){
    $domainSettingsObjectArray | ? { $_.domain -eq "Default" } | % { $_.SharePointSharing = "Blocked" }
    foreach($domain in $SPOSettings.sharingAllowedDomainList){

        if($domainSettingsObjectArray.Domain -contains $domain){
            $domainSettingsObjectArray | ? { $_.domain -eq $domain } | % { $_.SharePointSharing = "Allowed" }
        }else{
        $domainSettingsObject = [pscustomobject]@{
            tenantID                     = ""
            Domain                       = $domain
            GuestInvitations             = ""
            TeamsFederation              = ""
            SharePointSharing            = "Allowed"
            B2BCrossTenantAccessInbound  = ""
            B2BCrossTenantAccessOutbound = ""
            B2BDirectConnectInbound      = ""
            B2BDirectConnectOutbound     = ""
            CrossTenantSyncEnabled       = ""
            TrustSettings                = ""
            AutomaticRedemption          = ""
        }
    }
        $domainSettingsObjectArray += $domainSettingsObject
    }
}elseif(($SPOSettings.sharingCapability -ne "Disabled") -and ($SPOSettings.sharingDomainRestrictionMode -eq "none")){
    $domainSettingsObjectArray | ? { $_.domain -eq "Default" } | % { $_.SharePointSharing = "Allowed"}s 
}

##Get external domains from TeamsFederationConfiguration
$TeamsFederationSettings = Get-CsTenantFederationConfiguration

if (([string]$TeamsFederationSettings.alloweddomains -eq "AllowAllKnownDomains") -and ($TeamsFederationSettings.AllowFederatedUsers)) {
    $domainSettingsObjectArray | ? { $_.domain -eq "Default" } | % { $_.TeamsFederation = "Allowed" }
    foreach ($domain in [array]$TeamsFederationSettings.BlockedDomains.domain) {
        if ($domainSettingsObjectArray.Domain -contains $domain) {
            $domainSettingsObjectArray | ? { $_.domain -eq $domain } | % { $_.TeamsFederation = "Blocked" }
        }else {
            $domainSettingsObject = [pscustomobject]@{
                tenantID                     = ""
                Domain                       = $domain
                GuestInvitations             = "Org Default"
                TeamsFederation              = "Blocked"
                SharePointSharing            = ""
                B2BCrossTenantAccessInbound  = ""
                B2BCrossTenantAccessOutbound = ""
                B2BDirectConnectInbound      = ""
                B2BDirectConnectOutbound     = ""
                CrossTenantSyncEnabled       = ""
                TrustSettings                = ""
                AutomaticRedemption          = ""
            }
            $domainSettingsObjectArray += $domainSettingsObject
        }
    }
}else {
    $domainSettingsObjectArray | ? { $_.domain -eq "Default" } | % { $_.TeamsFederation = "Blocked" }
    foreach ($domain in [array]$TeamsFederationSettings.AllowedDomains.AllowedDomain.domain) {
        if ($domainSettingsObjectArray.Domain -contains $domain) {
            $domainSettingsObjectArray | ? { $_.domain -eq $domain } | % { $_.TeamsFederation = "Allowed" }
        }else {
            $domainSettingsObject = [pscustomobject]@{
                tenantID                     = ""
                Domain                       = $domain
                GuestInvitations             = "Org Default"
                TeamsFederation              = "Allowed"
                SharePointSharing            = ""
                B2BCrossTenantAccessInbound  = ""
                B2BCrossTenantAccessOutbound = ""
                B2BDirectConnectInbound      = ""
                B2BDirectConnectOutbound     = ""
                CrossTenantSyncEnabled       = ""
                TrustSettings                = ""
                AutomaticRedemption          = ""
            }
            $domainSettingsObjectArray += $domainSettingsObject
        }
    }
}



##Update Tenant IDs for all existing domains
foreach ($domain in $domainSettingsObjectArray) {
    Try {
        $tenantID = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/tenantRelationships/findTenantInformationByDomainName(domainName='$($domain.Domain)')").tenantID
        $domain.tenantID = $tenantID
    }
    Catch {
        $tenantID = "N/A"
    }
    $domain.tenantID = $tenantID
}


##Get Cross Tenant Policy Defaults
$DefaultCrossTenantPolicy = Get-MgPolicyCrossTenantAccessPolicyDefault

##Process B2B Collaboration Inbound
$B2BAppsPolicy = $DefaultCrossTenantPolicy.B2BCollaborationInbound.Applications.AccessType
$B2BUsersPolicy = $DefaultCrossTenantPolicy.B2BCollaborationInbound.UsersAndGroups.AccessType
if ($DefaultCrossTenantPolicy.B2BCollaborationInbound.Applications.Targets.target -eq "AllApplications") {
    $B2BAppsTarget = "AllApps"
}else {
    $B2BAppsTarget = "SelectedApps"
}

if ($DefaultCrossTenantPolicy.B2BCollaborationInbound.UsersAndGroups.Targets.target -eq "AllUsers") {
    $B2BUsersTarget = "AllUsers"
}else {
    $B2BUsersTarget = "SelectedUsers"
}

$CrossTenantAccessInbound = "$B2BAppsPolicy $B2BAppsTarget`n$B2BUsersPolicy $B2BUsersTarget"


##Process B2B Collaboration Outbound
$B2BAppsPolicy = $DefaultCrossTenantPolicy.B2BCollaborationOutbound.Applications.AccessType
$B2BUsersPolicy = $DefaultCrossTenantPolicy.B2BCollaborationOutbound.UsersAndGroups.AccessType
if ($DefaultCrossTenantPolicy.B2BCollaborationOutbound.Applications.Targets.target -eq "AllApplications") {
    $B2BAppsTarget = "AllApps"
}else {
    $B2BAppsTarget = "SelectedApps"
}

if ($DefaultCrossTenantPolicy.B2BCollaborationOutbound.UsersAndGroups.Targets.target -eq "AllUsers") {
    $B2BUsersTarget = "AllUsers"
}else {
    $B2BUsersTarget = "SelectedUsers"
}

$CrossTenantAccessOutbound = "$B2BAppsPolicy $B2BAppsTarget`n$B2BUsersPolicy $B2BUsersTarget"

##Update Default B2B Collaboration Inbound and Outbound
$domainSettingsObjectArray | ? { $_.domain -eq "Default" } | % { $_.B2BCrossTenantAccessInbound = "$CrossTenantAccessInbound" }
$domainSettingsObjectArray | ? { $_.domain -eq "Default" } | % { $_.B2BCrossTenantAccessOutbound = "$CrossTenantAccessOutbound" }

##Process Default B2B Direct Connect Inbound
$B2BAppsPolicy = $DefaultCrossTenantPolicy.B2BDirectConnectInbound.Applications.AccessType
$B2BUsersPolicy = $DefaultCrossTenantPolicy.B2BDirectConnectInbound.UsersAndGroups.AccessType
if ($DefaultCrossTenantPolicy.B2BDirectConnectInbound.Applications.Targets.target -eq "AllApplications") {
    $B2BAppsTarget = "AllApps"
}else {
    $B2BAppsTarget = "SelectedApps"
}

if ($DefaultCrossTenantPolicy.B2BDirectConnectInbound.UsersAndGroups.Targets.target -eq "AllUsers") {
    $B2BUsersTarget = "AllUsers"
}else {
    $B2BUsersTarget = "SelectedUsers"
}

$DirectConnectAccessInbound = "$B2BAppsPolicy $B2BAppsTarget`n$B2BUsersPolicy $B2BUsersTarget"

$B2BAppsPolicy = $DefaultCrossTenantPolicy.B2BDirectConnectOutbound.Applications.AccessType
$B2BUsersPolicy = $DefaultCrossTenantPolicy.B2BDirectConnectOutbound.UsersAndGroups.AccessType
if ($DefaultCrossTenantPolicy.B2BDirectConnectOutbound.Applications.Targets.target -eq "AllApplications") {
    $B2BAppsTarget = "AllApps"
}else {
    $B2BAppsTarget = "SelectedApps"
}

if ($DefaultCrossTenantPolicy.B2BDirectConnectOutbound.UsersAndGroups.Targets.target -eq "AllUsers") {
    $B2BUsersTarget = "AllUsers"
}else {
    $B2BUsersTarget = "SelectedUsers"
}

$DirectConnectAccessOutbound = "$B2BAppsPolicy $B2BAppsTarget`n$B2BUsersPolicy $B2BUsersTarget"

$domainSettingsObjectArray | ? { $_.domain -eq "Default" } | % { $_.B2BDirectConnectInbound = "$DirectConnectAccessInbound" }
$domainSettingsObjectArray | ? { $_.domain -eq "Default" } | % { $_.B2BDirectConnectOutbound = "$DirectConnectAccessOutbound" }

#Process Default Cross Tenant Sync Inbound and Outbound
$domainSettingsObjectArray | ? { $_.domain -eq "Default" } | % { $_.CrossTenantSyncEnabled = "N/A" }


##Process Default Trust Settings and Automatic Redemption
$domainSettingsObjectArray | ? { $_.domain -eq "Default" } | % { $_.TrustSettings = "N/A" }
$domainSettingsObjectArray | ? { $_.domain -eq "Default" } | % { $_.AutomaticRedemption = "N/A" }

##Get Cross Tenant Policies for Partner Domains
[array]$CrossTenantPartnerPolicies = Get-MgPolicyCrossTenantAccessPolicyPartner

##Process Partner Domains
foreach ($CrossTenantPartnerPolicy in $CrossTenantPartnerPolicies) {

    ##Get Tenant Domain for Partner
    $tenantDomain = (Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/tenantRelationships/findTenantInformationByTenantId(TenantID='$($CrossTenantPartnerPolicy.TenantID)')").defaultDomainName

    ##Process B2B Collaboration Inbound
    $B2BAppsPolicy = $CrossTenantPartnerPolicy.B2BCollaborationInbound.Applications.AccessType
    $B2BUsersPolicy = $CrossTenantPartnerPolicy.B2BCollaborationInbound.UsersAndGroups.AccessType
    if ($CrossTenantPartnerPolicy.B2BCollaborationInbound.Applications.Targets.target -eq "AllApplications") {
        $B2BAppsTarget = "AllApps"
    }else {
        $B2BAppsTarget = "SelectedApps"
    }

    if ($CrossTenantPartnerPolicy.B2BCollaborationInbound.UsersAndGroups.Targets.target -eq "AllUsers") {
        $B2BUsersTarget = "AllUsers"
    }else {
        $B2BUsersTarget = "SelectedUsers"
    }

    if($B2BAppsPolicy){
    $CrossTenantAccessInbound = "$B2BAppsPolicy $B2BAppsTarget`n$B2BUsersPolicy $B2BUsersTarget"
    }else{
        $CrossTenantAccessInbound = "Org Default"
    }

    ##Process B2B Collaboration Outbound
    $B2BAppsPolicy = $CrossTenantPartnerPolicy.B2BCollaborationOutbound.Applications.AccessType
    $B2BUsersPolicy = $CrossTenantPartnerPolicy.B2BCollaborationOutbound.UsersAndGroups.AccessType
    if ($CrossTenantPartnerPolicy.B2BCollaborationOutbound.Applications.Targets.target -eq "AllApplications") {
        $B2BAppsTarget = "AllApps"
    }else {
        $B2BAppsTarget = "SelectedApps"
    }

    if ($CrossTenantPartnerPolicy.B2BCollaborationOutbound.UsersAndGroups.Targets.target -eq "AllUsers") {
        $B2BUsersTarget = "AllUsers"
    }else {
        $B2BUsersTarget = "SelectedUsers"
    }

    if($B2BAppsPolicy){
        $CrossTenantAccessOutbound = "$B2BAppsPolicy $B2BAppsTarget`n$B2BUsersPolicy $B2BUsersTarget"
        }else{
            $CrossTenantAccessOutbound = "Org Default"
        }
    

        
    ##Process B2B Direct Connect Inbound
    $B2BAppsPolicy = $CrossTenantPartnerPolicy.B2BDirectConnectInbound.Applications.AccessType
    $B2BUsersPolicy = $CrossTenantPartnerPolicy.B2BDirectConnectInbound.UsersAndGroups.AccessType
    if ($CrossTenantPartnerPolicy.B2BDirectConnectInbound.Applications.Targets.target -eq "AllApplications") {
        $B2BAppsTarget = "AllApps"
    }else {
        $B2BAppsTarget = "SelectedApps"
    }

    if ($CrossTenantPartnerPolicy.B2BDirectConnectInbound.UsersAndGroups.Targets.target -eq "AllUsers") {
        $B2BUsersTarget = "AllUsers"
    }else {
        $B2BUsersTarget = "SelectedUsers"
    }

    if($B2BAppsPolicy){
        $DirectConnectAccessInbound = "$B2BAppsPolicy $B2BAppsTarget`n$B2BUsersPolicy $B2BUsersTarget"
        }else{
            $DirectConnectAccessInbound = "Org Default"
        }
    

    ##Process B2B Direct Connect Outbound
    $B2BAppsPolicy = $CrossTenantPartnerPolicy.B2BDirectConnectOutbound.Applications.AccessType
    $B2BUsersPolicy = $CrossTenantPartnerPolicy.B2BDirectConnectOutbound.UsersAndGroups.AccessType
    if ($CrossTenantPartnerPolicy.B2BDirectConnectOutbound.Applications.Targets.target -eq "AllApplications") {
        $B2BAppsTarget = "AllApps"
    }else {
        $B2BAppsTarget = "SelectedApps"
    }

    if ($CrossTenantPartnerPolicy.B2BDirectConnectOutbound.UsersAndGroups.Targets.target -eq "AllUsers") {
        $B2BUsersTarget = "AllUsers"
    }else {
        $B2BUsersTarget = "SelectedUsers"
    }

        if($B2BAppsPolicy){
        $DirectConnectAccessOutbound = "$B2BAppsPolicy $B2BAppsTarget`n$B2BUsersPolicy $B2BUsersTarget"
        }else{
            $DirectConnectAccessOutbound = "Org Default"
        }

    ##Process Cross tenant sync inbound
    
    Try {
        if((Get-MgPolicyCrossTenantAccessPolicyPartnerIdentitySynchronization -CrossTenantAccessPolicyConfigurationPartnerTenantId $CrossTenantPartnerPolicy.TenantId -erroraction silentlycontinue).UserSyncInbound.issyncallowed){
            $SyncEnabled = "Enabled"
        }else{
            $SyncEnabled = "Disabled"
        }
    }
    catch {
        $syncEnabled = "Disabled"
    }

    ##Process Trusted Settings
    $TrustedSettings = ""

    if ($crosstenantpartnerPolicy.inboundtrust.IsCompliantDeviceAccepted) {
        $TrustedSettings += "Compliant Devices`n"
    }

    if ($CrossTenantPartnerPolicy.inboundtrust.IsMfaAccepted) {
        $TrustedSettings += "Domain Joined Devices`n"
    }

    if ($crosstenantpartnerPolicy.InboundTrust.IsIntuneManagedDeviceAccepted) {
        $TrustedSettings += "Intune Managed Devices`n"
    }

    if (!$trustedSettings) {
        $TrustedSettings = "None"
    }

    ##Process Automatic Redemption
    if ($crosstenantpartnerPolicy.AutomaticUserConsentSettings.InboundAllowed) {
        $AutomaticRedemption = "Enabled"
    }else {
        $AutomaticRedemption = "Disabled"
    }

    ##Chek if the domain already exists in the array and update
    if ($domainSettingsObjectArray.Domain -contains $tenantDomain) {
        ##Update existing domain
        $domainSettingsObjectArray | ? { $_.domain -eq $tenantDomain } | % { $_.B2BCrossTenantAccessInbound = "$CrossTenantAccessInbound" }
        $domainSettingsObjectArray | ? { $_.domain -eq $tenantDomain } | % { $_.B2BCrossTenantAccessOutbound = "$CrossTenantAccessOutbound" }
        $domainSettingsObjectArray | ? { $_.domain -eq $tenantDomain } | % { $_.B2BDirectConnectInbound = "$DirectConnectAccessInbound" }
        $domainSettingsObjectArray | ? { $_.domain -eq $tenantDomain } | % { $_.B2BDirectConnectOutbound = "$DirectConnectAccessOutbound" }
        $domainSettingsObjectArray | ? { $_.domain -eq $tenantDomain } | % { $_.CrossTenantSyncEnabled = "$SyncEnabled" }
        $domainSettingsObjectArray | ? { $_.domain -eq $tenantDomain } | % { $_.TrustSettings = "$TrustedSettings" }
        $domainSettingsObjectArray | ? { $_.domain -eq $tenantDomain } | % { $_.AutomaticRedemption = "$AutomaticRedemption" }
    }else {
        ##Add new domain to array
        $domainSettingsObject = [pscustomobject]@{
            tenantID                     = $CrossTenantPartnerPolicy.TenantID
            Domain                       = $tenantDomain
            GuestInvitations             = "Org Default"
            TeamsFederation              = "Org Default"
            SharePointSharing            = "Org Default"
            B2BCrossTenantAccessInbound  = "$CrossTenantAccessInbound"
            B2BCrossTenantAccessOutbound = "$CrossTenantAccessOutbound"
            B2BDirectConnectInbound      = "$DirectConnectAccessInbound"
            B2BDirectConnectOutbound     = "$DirectConnectAccessOutbound"
            CrossTenantSyncEnabled       = "$SyncEnabled"
            TrustSettings                = "$TrustedSettings"
            AutomaticRedemption          = "$AutomaticRedemption"
        }
        $domainSettingsObjectArray += $domainSettingsObject
    }

    If (($domainSettingsObjectArray.tenantID -contains "$TenantID") -and ($TenantID -ne "N/A")) {
        ##Update existing domain
        $domainSettingsObjectArray | ? { $_.tenantID -eq $TenantID } | % { $_.B2BCrossTenantAccessInbound = "$CrossTenantAccessInbound" }
        $domainSettingsObjectArray | ? { $_.tenantID -eq $TenantID } | % { $_.B2BCrossTenantAccessOutbound = "$CrossTenantAccessOutbound" }
        $domainSettingsObjectArray | ? { $_.tenantID -eq $TenantID } | % { $_.B2BDirectConnectInbound = "$DirectConnectAccessInbound" }
        $domainSettingsObjectArray | ? { $_.tenantID -eq $TenantID } | % { $_.B2BDirectConnectOutbound = "$DirectConnectAccessOutbound" }
        $domainSettingsObjectArray | ? { $_.tenantID -eq $TenantID } | % { $_.CrossTenantSyncEnabled = "$SyncEnabled" }
        $domainSettingsObjectArray | ? { $_.tenantID -eq $TenantID } | % { $_.TrustSettings = "$TrustedSettings" }  
        $domainSettingsObjectArray | ? { $_.tenantID -eq $TenantID } | % { $_.AutomaticRedemption = "$AutomaticRedemption" }      
    }
}

##Update Defaults for remaining domains
foreach ($domain in $domainSettingsObjectArray) {
    if (!$domain.teamsfederation) {
        $domain.teamsfederation = "Org Default"
    }
    if (!$domain.GuestInvitations) {
        $domain.GuestInvitations = "Org Default"
    }
    if (!$domain.SharePointSharing) {
        $domain.SharePointSharing = "Org Default"
    }
    if (!$domain.B2BCrossTenantAccessInbound) {
        $domain.B2BCrossTenantAccessInbound = "Org Default"
    }
    if (!$domain.B2BCrossTenantAccessOutbound) {
        $domain.B2BCrossTenantAccessOutbound = "Org Default"
    }
    if (!$domain.B2BDirectConnectInbound) {
        $domain.B2BDirectConnectInbound = "Org Default"
    }
    if (!$domain.B2BDirectConnectOutbound) {
        $domain.B2BDirectConnectOutbound = "Org Default"
    }
    if (!$domain.CrossTenantSyncEnabled) {
        $domain.CrossTenantSyncEnabled = "N/A"
    }
    if (!$domain.TrustSettings) {
        $domain.TrustSettings = "N/A"
    }
    if (!$domain.AutomaticRedemption) {
        $domain.AutomaticRedemption = "N/A"
    }
}

##Output to screen and file
$domainSettingsObjectArray | fl
$domainSettingsObjectArray | export-csv -path c:\temp\ExternalDomainConfiguration.csv -NoTypeInformation