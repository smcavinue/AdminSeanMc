<##Author: Sean McAvinue
##Details: Graph / PowerShell Script to assess a Microsoft 365 tenant for migration of Exchange, Teams, SharePoint and OneDrive, 
##          Please fully read and test any scripts before running in your production environment!
        .SYNOPSIS
        Provides a detailed report of Conditional Access configuration across users and services

        .DESCRIPTION
        Gathers information using Microsoft Graph API For Conditional Access and exports as a report in Excel

        .PARAMETER ClientID
        Required - Application (Client) ID of the App Registration

        .PARAMETER TenantID
        Required - Directory (Tenant) ID of the Azure AD Tenant

        .PARAMETER certificateThumbprint
        Required - Thumbprint of the certificate generated from the prepare-ConditionalAccessDetailReport.ps1 script    

        .Notes
        For similar scripts check out the links below
        
            Blog: https://seanmcavinue.net
            GitHub: https://github.com/smcavinue
            Twitter: @Sean_McAvinue
            Linkedin: https://www.linkedin.com/in/sean-mcavinue-4a058874/


    #>
#Requires -modules msal.ps, importexcel
Param(
    [parameter(Mandatory = $true,
        ParameterSetName = 'Certificate')]
    [Parameter(Mandatory = $true,
        ParameterSetName = 'Secret')]
    [Parameter(Mandatory = $true,
        ParameterSetName = 'Delegated')]
    $clientId,
    [parameter(Mandatory = $true,
        ParameterSetName = 'Certificate')]
    [Parameter(Mandatory = $true,
        ParameterSetName = 'Secret')]
    [Parameter(Mandatory = $true,
        ParameterSetName = 'Delegated')]
    $tenantId,
    [parameter(Mandatory = $true,
        ParameterSetName = 'Certificate')]
    $certificateThumbprint,
    [parameter(Mandatory = $true,
        ParameterSetName = 'Secret')]
    $Secret,
    [Parameter(Mandatory = $false)]
    [Switch]$ShowGraphCalls
)

function UpdateProgress {
    Write-Progress -Activity "Conditional Access Assessment in Progress" -Status "Processing Task $ProgressTracker of $($TotalProgressTasks): $ProgressStatus" -PercentComplete (($ProgressTracker / $TotalProgressTasks) * 100)
}

function RunQueryandEnumerateResults {
    <#
    .SYNOPSIS
    Runs Graph Query and if there a
    re any additional pages, parses them and appends to a single variable

    
    #>
    if ($ShowGraphCalls) {
        write-host $apiuri
    }
    #Run Graph Query
    Try {
        $Results = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.accesstoken)" } -Uri $apiUri -Method Get)
    }
    catch {
        write-host $PassthroughError -ForegroundColor red
    }
    #Output Results for Debugging
    #write-host $results

    #Begin populating results
    [array]$ResultsValue = $Results.value

    #If there is a next page, query the next page until there are no more pages and append results to existing set
    if ($results."@odata.nextLink" -ne $null) {
        $NextPageUri = $results."@odata.nextLink"
        ##While there is a next page, query it and loop, append results
        While ($NextPageUri -ne $null) {
            $NextPageRequest = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.accesstoken)" } -Uri $NextPageURI -Method Get)
            $NxtPageData = $NextPageRequest.Value
            $NextPageUri = $NextPageRequest."@odata.nextLink"
            $ResultsValue = $ResultsValue + $NxtPageData
        }
    }

    ##Return completed results
    return $ResultsValue

    
}

function ExpandGroup {
    Param(
        [parameter(Mandatory = $true)]
        $GroupDetails
    )
    $global:grouplist += $GroupDetails 
    $GroupSplit = $groupDetails -split " - *"
    $NestedGroupID = $GroupSplit[($groupsplit.count - 1)]
    $global:grouplistPrevious = $global:grouplist

    $apiuri = "https://graph.microsoft.com/v1.0/groups/$($NestedGroupID)/members"
    #write-host $apiuri
    [array]$NestedMembers = RunQueryandEnumerateResults
    foreach ($Nestedmember in $NestedMembers) {

        if ($Nestedmember.'@odata.type' -eq "#microsoft.graph.user") {

            $Row = New-Object psobject -Property @{
                "User"                                     = $Nestedmember.userPrincipalName
                "User Excluded"                            = "No"
                "Apps"                                     = ($policy.conditions.applications.includeApplications -join ';')
                "Policy"                                   = $Policy.displayname
                "Policy State"                             = $Policy.state
                "Inherited from Group"                     = ($global:grouplist -join ';')
                "Exclusion inherited from Group"           = "No"
                "Inherited from Role Assignment"           = "No"
                "Exclusion inherited from Role Assignment" = "No"
            }
            $global:UserOutput += $row
        }
        elseif ($Nestedmember.'@odata.type' -eq "#microsoft.graph.group") {
           
            If ($global:grouplist -contains "$($Nestedmember.displayname) - $($Nestedmember.id)") {
                write-host "Nested Group loop detected for group $($Nestedmember.displayname) in path $($global:grouplist), consider removing this loop" -ForegroundColor Yellow
            }
            else {
                ExpandGroup -GroupDetails "$($Nestedmember.displayname) - $($Nestedmember.id)"
            }
            $global:grouplist = $global:grouplistPrevious
        }
    }
}

function ExpandExclusionGroup {
    Param(
        [parameter(Mandatory = $true)]
        $GroupDetails
    )

    $global:grouplist += $GroupDetails 
    $GroupSplit = $groupDetails -split " - *"
    $NestedGroupID = $GroupSplit[($groupsplit.count - 1)]
    $global:grouplistPrevious = $global:grouplist

    $apiuri = "https://graph.microsoft.com/v1.0/groups/$($NestedGroupID)/members"
    #write-host $apiuri
    [array]$NestedMembers = RunQueryandEnumerateResults
    foreach ($Nestedmember in $NestedMembers) {

        if ($Nestedmember.'@odata.type' -eq "#microsoft.graph.user") {

            $ExcludedUPN = $Nestedmember.userprincipalname

            [array]$PoliciesImpacted = $global:UserOutput | ? { ($_.user -eq $ExcludedUPN) -and ($_.Policy -eq $Policy.displayname) }
        
            foreach ($PolicyImpacted in $PoliciesImpacted) {
                        
                if ($PolicyImpacted.'Exclusion inherited from Group' -eq "No") {
                    $PolicyImpacted.'User Excluded' = "Yes"
                    $PolicyImpacted.'Exclusion inherited from Group' = ($global:grouplist -join ';')
                }
                else {
                    $PolicyImpacted.'Exclusion inherited from Group' = ($PolicyImpacted.'Exclusion inherited from Group' + " & " + ($global:grouplist -join ';'))
                }
        
            }
        }
        elseif ($Nestedmember.'@odata.type' -eq "#microsoft.graph.group") {
           
            If ($global:grouplist -contains "$($Nestedmember.displayname) - $($Nestedmember.id)") {
                write-host "Nested Group loop detected for group $($Nestedmember.displayname) in path $($global:grouplist), consider removing this loop" -ForegroundColor Yellow
            }
            else {
                ExpandExclusionGroup -GroupDetails "$($Nestedmember.displayname) - $($Nestedmember.id)"
            }
            $global:grouplist = $global:grouplistPrevious
        }
    }
}

##Report File Name
$Filename = "ConditionalAccessAssessment-$((get-date).tostring().replace('/','').replace(':','')).xlsx"
##File Location
$FilePath = "C:\temp"
Try {
    if (!(test-path -Path $FilePath)) {
        New-Item -Path $FilePath -ItemType Directory
    }
}
catch {
    write-host "Could not create folder at c:\temp - check you have appropriate permissions" -ForegroundColor red
    exit
}


$ProgressTracker = 1
$TotalProgressTasks = 1
$ProgressStatus = "Obtaining Graph Token"
UpdateProgress

##Attempt to get an Access Token
Try {
    If ($certificateThumbprint) {
        $CertificatePath = "cert:\currentuser\my\$CertificateThumbprint"
        $Certificate = Get-Item $certificatePath
        $Token = Get-MsalToken -ClientId $ClientId -TenantId $TenantId -ClientCertificate $Certificate
    }
    elseif ($secret) {
        $Token = Get-MsalToken -ClientId $ClientId -TenantId $TenantId -ClientSecret ($Secret | ConvertTo-SecureString -AsPlainText -Force) -ForceRefresh
    }
    else {
        $Token = Get-MsalToken -ClientId $ClientId -TenantId $TenantId -RedirectUri "https://localhost" -ForceRefresh 
    }
}
Catch {
    write-host "Unable to acquire access token, check the parameters are correct`n$($Error[0])"
    exit
}

##Defing Group tracking variables
$global:grouplist = @()
$global:grouplistPrevious = @()

##Update progress
$ProgressTracker = 2
$TotalProgressTasks = 11
$ProgressStatus = "Fetching Conditional Access Policies"
UpdateProgress
$ProgressTracker++

##Get Conditional Access Policies
$apiURI = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"
[array]$ConditionalAccessPolicies = RunQueryandEnumerateResults

##Update Progress
$TotalProgressTasks = (11 + $ConditionalAccessPolicies.count)
$ProgressStatus = "Fetching Directory Role Templates"
UpdateProgress
$ProgressTracker++

##Get Directory Roles
$apiURI = "https://graph.microsoft.com/beta/directoryRoleTemplates"
[array]$DirectoryRoleTemplates = RunQueryandEnumerateResults

##Update Progress
$ProgressStatus = "Fetching Directory Roles"
UpdateProgress
$ProgressTracker++

##Get enabled Directory Roles
$apiURI = "https://graph.microsoft.com/v1.0/directoryRoles"
[array]$EnabledDirectoryRoles = RunQueryandEnumerateResults

##Update Progress
$ProgressStatus = "Fetching Trusted Locations"
UpdateProgress
$ProgressTracker++

##Get Trusted Locations
$apiURI = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations"
[array]$NamedLocations = RunQueryandEnumerateResults

##Update Progress
$ProgressStatus = "Fetching Users"
UpdateProgress
$ProgressTracker++

##List All Tenant Users
$apiuri = "https://graph.microsoft.com/v1.0/users?`$select=id,displayname,userprincipalname"
$users = RunQueryandEnumerateResults

##Update Progress
$ProgressStatus = "Fetching Groups"
UpdateProgress
$ProgressTracker++

##List all Tenant Groups
$apiuri = "https://graph.microsoft.com/v1.0/groups"
[array]$Groups = RunQueryandEnumerateResults

##Update Progress
$ProgressStatus = "Fetching Service Principals"
UpdateProgress
$ProgressTracker++

##List All Azure AD Service Principals
$apiURI = "https://graph.microsoft.com/beta/servicePrincipals"
[array]$AADApps = RunQueryandEnumerateResults

##Update Progress
$ProgressStatus = "Updating Conditional Access Export"
UpdateProgress
$ProgressTracker++

##Tidy GUIDs to names
$ConditionalAccessPoliciesJSON = $ConditionalAccessPolicies | ConvertTo-Json -Depth 5
if ($ConditionalAccessPoliciesJSON -ne $null) {
    ##TidyUsers
    foreach ($User in $Users) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($user.id, ("$($user.displayname) - $($user.userPrincipalName)"))
    }

    ##Tidy Groups
    foreach ($Group in $Groups) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($group.id, ("$($group.displayname) - $($group.id)"))
    }

    ##Tidy Roles
    foreach ($DirectoryRoleTemplate in $DirectoryRoleTemplates) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($DirectoryRoleTemplate.Id, $DirectoryRoleTemplate.displayname)
    }

    ##Tidy Apps
    foreach ($AADApp in $AADApps) {
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($AADApp.appid, $AADApp.displayname)
        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($AADApp.id, $AADApp.displayname)
    }

    ##Tidy Locations
    foreach ($NamedLocation in $NamedLocations) {

        $ConditionalAccessPoliciesJSON = $ConditionalAccessPoliciesJSON.Replace($NamedLocation.id, $NamedLocation.displayname)
    }


    $ConditionalAccessPolicies = $ConditionalAccessPoliciesJSON | ConvertFrom-Json


    $CAOutput = @()
    $CAHeadings = @(
        "createdDateTime",
        "modifiedDateTime",
        "state",
        "Conditions.users.includeusers",
        "Conditions.users.excludeusers",
        "Conditions.users.includegroups",
        "Conditions.users.excludegroups",
        "Conditions.users.includeroles",
        "Conditions.users.excluderoles",
        "Conditions.clientApplications.includeServicePrincipals",
        "Conditions.clientApplications.excludeServicePrincipals",
        "Conditions.applications.includeApplications",
        "Conditions.applications.excludeApplications",
        "Conditions.applications.includeUserActions",
        "Conditions.applications.includeAuthenticationContextClassReferences",
        "Conditions.userRiskLevels",
        "Conditions.signInRiskLevels",
        "Conditions.platforms.includePlatforms",
        "Conditions.platforms.excludePlatforms",
        "Conditions.locations.includLocations",
        "Conditions.locations.excludeLocations",
        "Conditions.clientAppTypes",
        "Conditions.devices.deviceFilter.mode",
        "Conditions.devices.deviceFilter.rule",
        "GrantControls.operator",
        "grantcontrols.builtInControls",
        "grantcontrols.customAuthenticationFactors",
        "grantcontrols.termsOfUse",
        "SessionControls.disableResilienceDefaults",
        "SessionControls.applicationEnforcedRestrictions",
        "SessionControls.persistentBrowser",
        "SessionControls.cloudAppSecurity",
        "SessionControls.signInFrequency"

    )
    foreach ($Policy in $ConditionalAccessPolicies) {
        $CAObject = @{
            "Policy Name"                                                         = $policy.displayName
            "createdDateTime"                                                     = $Policy.createdDateTime
            "modifiedDateTime"                                                    = $Policy.modifiedDateTime
            "state"                                                               = $Policy.state
            "Conditions.users.includeusers"                                       = $Policy.Conditions.users.includeusers -join ";"
            "Conditions.users.excludeusers"                                       = $Policy.Conditions.users.excludeusers -join ';'
            "Conditions.users.includegroups"                                      = $Policy.Conditions.users.includegroups -join ';'
            "Conditions.users.excludegroups"                                      = $Policy.Conditions.users.excludegroups -join ';'
            "Conditions.users.includeroles"                                       = $Policy.Conditions.users.includeroles -join ';'
            "Conditions.users.excluderoles"                                       = $Policy.Conditions.users.excluderoles -join ';'
            "Conditions.clientApplications.includeServicePrincipals"              = $Policy.Conditions.clientApplications.includeServicePrincipals -join ';'
            "Conditions.clientApplications.excludeServicePrincipals"              = $Policy.Conditions.clientApplications.excludeServicePrincipals -join ';'
            "Conditions.applications.includeApplications"                         = $Policy.Conditions.applications.includeApplications -join ';'
            "Conditions.applications.excludeApplications"                         = $Policy.Conditions.applications.excludeApplications -join ';'
            "Conditions.applications.includeUserActions"                          = $Policy.Conditions.applications.includeUserActions -join ';'
            "Conditions.applications.includeAuthenticationContextClassReferences" = $Policy.Conditions.applications.includeAuthenticationContextClassReferences -join ';'
            "Conditions.userRiskLevels"                                           = $Policy.Conditions.userRiskLevels -join ';'
            "Conditions.signInRiskLevels"                                         = $Policy.Conditions.signInRiskLevels -join ';'
            "Conditions.platforms.includePlatforms"                               = $Policy.Conditions.platforms.includePlatforms -join ';'
            "Conditions.platforms.excludePlatforms"                               = $Policy.Conditions.platforms.excludePlatforms -join ';'
            "Conditions.locations.includLocations"                                = $Policy.Conditions.locations.includLocations -join ';'
            "Conditions.locations.excludeLocations"                               = $Policy.Conditions.locations.excludeLocations -join ';'
            "Conditions.clientAppTypes"                                           = $Policy.Conditions.clientAppTypes -join ';'
            "Conditions.devices.deviceFilter.mode"                                = $Policy.Conditions.devices.deviceFilter.mode -join ';'
            "Conditions.devices.deviceFilter.rule"                                = $Policy.Conditions.devices.deviceFilter.rule -join ';'
            "GrantControls.operator"                                              = $Policy.GrantControls.operator -join ';'
            "grantcontrols.builtInControls"                                       = $Policy.grantcontrols.builtInControls -join ';'
            "grantcontrols.customAuthenticationFactors"                           = $Policy.grantcontrols.customAuthenticationFactors -join ';'
            "grantcontrols.termsOfUse"                                            = $Policy.grantcontrols.termsOfUse -join ';'
            "SessionControls.disableResilienceDefaults"                           = $Policy.SessionControls.disableResilienceDefaults -join ';'
            "SessionControls.applicationEnforcedRestrictions"                     = $Policy.SessionControls.applicationEnforcedRestrictions -join ';'
            "SessionControls.persistentBrowser"                                   = $Policy.SessionControls.persistentBrowser -join ';'
            "SessionControls.cloudAppSecurity"                                    = $Policy.SessionControls.cloudAppSecurity -join ';'
            "SessionControls.signInFrequency"                                     = $Policy.SessionControls.signInFrequency -join ';'

        }
        [array]$CAColumns += [PSCustomObject]$CAObject
        
    }
    Foreach ($Heading in $CAHeadings) {
        $Row = $null
        $Row = New-Object psobject -Property @{
            PolicyName = $Heading
        }
    
        foreach ($CAPolicy in $ConditionalAccessPolicies) {
            $Nestingcheck = ($Heading.split('.').count)

            if ($Nestingcheck -eq 1) {
                $Row | Add-Member -MemberType NoteProperty -Name $CAPolicy.displayname -Value $CAPolicy.$Heading -Force
            }
            elseif ($Nestingcheck -eq 2) {
                $SplitHeading = $Heading.split('.')
                $Row | Add-Member -MemberType NoteProperty -Name $CAPolicy.displayname -Value ($CAPolicy.($SplitHeading[0].ToString()).($SplitHeading[1].ToString()) -join ';' )-Force
            }
            elseif ($Nestingcheck -eq 3) {
                $SplitHeading = $Heading.split('.')
                $Row | Add-Member -MemberType NoteProperty -Name $CAPolicy.displayname -Value ($CAPolicy.($SplitHeading[0].ToString()).($SplitHeading[1].ToString()).($SplitHeading[2].ToString()) -join ';' )-Force
            }
            elseif ($Nestingcheck -eq 4) {
                $SplitHeading = $Heading.split('.')
                $Row | Add-Member -MemberType NoteProperty -Name $CAPolicy.displayname -Value ($CAPolicy.($SplitHeading[0].ToString()).($SplitHeading[1].ToString()).($SplitHeading[2].ToString()).($SplitHeading[3].ToString()) -join ';' )-Force       
            }
        }

        $CAOutput += $Row
        

    }
    $CAOutput | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Conditional Access by Column" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    $CAColumns | Select "Policy Name", "createdDateTime", "modifiedDateTime", "state", "Conditions.users.includeusers", "Conditions.users.excludeusers", "Conditions.users.includegroups", "Conditions.users.excludegroups", "Conditions.users.includeroles", "Conditions.users.excluderoles", "Conditions.clientApplications.includeServicePrincipals", "Conditions.clientApplications.excludeServicePrincipals", "Conditions.applications.includeApplications", "Conditions.applications.excludeApplications", "Conditions.applications.includeUserActions", "Conditions.applications.includeAuthenticationContextClassReferences", "Conditions.userRiskLevels", "Conditions.signInRiskLevels", "Conditions.platforms.includePlatforms", "Conditions.platforms.excludePlatforms", "Conditions.locations.includLocations", "Conditions.locations.excludeLocations", "Conditions.clientAppTypes", "Conditions.devices.deviceFilter.mode", "Conditions.devices.deviceFilter.rule", "GrantControls.operator", "grantcontrols.builtInControls", "grantcontrols.customAuthenticationFactors", "grantcontrols.termsOfUse", "SessionControls.disableResilienceDefaults", "SessionControls.applicationEnforcedRestrictions", "SessionControls.persistentBrowser", "SessionControls.cloudAppSecurity", "SessionControls.signInFrequency" | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "Conditional Access by Row" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
}

##Define output variable
$global:UserOutput = @()
foreach ($policy in $ConditionalAccessPolicies) {
    ##Update Progress
    $ProgressStatus = "Processing policy: $($policy.displayName)"
    UpdateProgress
    $ProgressTracker++
    $ProgressStatus = "Processing policy: $($policy.displayName) - Direct user assignments"
    UpdateProgress
    ##Assess direct assigned users
    foreach ($AssignedObject in ($policy.conditions.users.includeusers)) {
        If (($policy.conditions.users.includeusers -ne "All") -and ($policy.conditions.users.includeusers -eq "GuestsOrExternalUsers")) {  
            $Row = New-Object psobject -Property @{
                "User"                                     = $AssignedObject
                "User Excluded"                            = "No"
                "Apps"                                     = ($policy.conditions.applications.includeApplications -join ';')
                "Policy"                                   = $Policy.displayname
                "Policy State"                             = $Policy.state
                "Inherited from Group"                     = "Direct"
                "Exclusion inherited from Group"           = "No"
                "Inherited from Role Assignment"           = "No"
                "Exclusion inherited from Role Assignment" = "No"
            }

            $global:UserOutput += $row
        }

        ##Process policies targeted to all users
        If ($policy.conditions.users.includeusers -eq "All") {
        
            foreach ($user in $users) {
            
                $Row = New-Object psobject -Property @{
                    "User"                                     = $user.userPrincipalName
                    "User Excluded"                            = "No"
                    "Apps"                                     = ($policy.conditions.applications.includeApplications -join ';')
                    "Policy"                                   = $Policy.displayname
                    "Policy State"                             = $Policy.state
                    "Inherited from Group"                     = "All Users"
                    "Exclusion inherited from Group"           = "No"
                    "Inherited from Role Assignment"           = "No"
                    "Exclusion inherited from Role Assignment" = "No"
                }
                $global:UserOutput += $row

            }
        }
        ##Process policies targeted to all guests or external users
        If ($policy.conditions.users.includeusers -eq "GuestsOrExternalUsers") {
        
            foreach ($user in ($users | ? { $_.userprincipalname -like "*#EXT#@*" })) {

                $Row = New-Object psobject -Property @{
                    "User"                                     = $user.userPrincipalName
                    "User Excluded"                            = "No"
                    "Apps"                                     = ($policy.conditions.applications.includeApplications -join ';')
                    "Policy"                                   = $Policy.displayname
                    "Policy State"                             = $Policy.state
                    "Inherited from Group"                     = "All Guest or External Users"
                    "Exclusion inherited from Group"           = "No"
                    "Inherited from Role Assignment"           = "No"
                    "Exclusion inherited from Role Assignment" = "No"
                }
                $global:UserOutput += $row
            }
        }
    }
    ##Update Progress
    $ProgressStatus = "Processing policy: $($policy.displayName) - Group assignments"
    UpdateProgress
    ##Assess Group assigned users
    foreach ($AssignedGroups in ($policy.conditions.users.includegroups)) {

        foreach ($Group in $AssignedGroups) {
            $GroupSplit = $group -split " - *"
            $GroupID = $GroupSplit[($groupsplit.count - 1)]
            $apiuri = "https://graph.microsoft.com/v1.0/groups/$GroupID/members"
            $PassthroughError = "Error getting members for $Group - Group has possibly been deleted and should be removed from the policy $($policy.displayName)"
            [array]$Members = RunQueryandEnumerateResults
           
            ##Process group members
            foreach ($member in $Members) {

                if ($member.'@odata.type' -eq "#microsoft.graph.user") {

                    $Row = New-Object psobject -Property @{
                        "User"                                     = $member.userPrincipalName
                        "User Excluded"                            = "No"
                        "Apps"                                     = ($policy.conditions.applications.includeApplications -join ';')
                        "Policy"                                   = $Policy.displayname
                        "Policy State"                             = $Policy.state
                        "Inherited from Group"                     = $Group
                        "Exclusion inherited from Group"           = "No"
                        "Inherited from Role Assignment"           = "No"
                        "Exclusion inherited from Role Assignment" = "No"
                    }
                    $global:UserOutput += $row
                }
                elseif ($member.'@odata.type' -eq "#microsoft.graph.group") {
                    ##Expand nested group members
                    $global:grouplist += $Group
                    ExpandGroup -GroupDetails "$($member.displayName) - $($member.id)"
                    $global:grouplist = @()
                    $global:grouplistPrevious = @()
                }
            }
        }
    }
    ##Update Progress
    $ProgressStatus = "Processing policy: $($policy.displayName) - Role assignments"
    UpdateProgress
    ##Assess Role assigned users
    foreach ($role in $policy.conditions.users.includeroles) {
        $roleid = $null
        $roleid = ($EnabledDirectoryRoles | ? { $_.displayname -eq $role }).id
        if ($roleid) {
            $apiURI = "https://graph.microsoft.com/v1.0/directoryRoles/roleTemplateId=$roleid/members"
            [array]$Rolemembers = RunQueryandEnumerateResults

            foreach ($Rolemember in $Rolemembers) {

                $Row = New-Object psobject -Property @{
                    "User"                                     = $Rolemember.userPrincipalName
                    "User Excluded"                            = "No"
                    "Apps"                                     = ($policy.conditions.applications.includeApplications -join ';')
                    "Policy"                                   = $Policy.displayname
                    "Policy State"                             = $Policy.state
                    "Inherited from Group"                     = "No"
                    "Exclusion inherited from Group"           = "No"
                    "Inherited from Role Assignment"           = $role
                    "Exclusion inherited from Role Assignment" = "No"
                }
                $global:UserOutput += $row

            }

        }
    }

    ##Assess direct excluded users
    $ProgressStatus = "Processing policy: $($policy.displayName) - Direct user exclusions"
    UpdateProgress
    foreach ($ExcludedObject in ($policy.conditions.users.Excludeusers)) {

        If ($ExcludedObject -ne "GuestsOrExternalUsers") {  
            $ExcludedUPN = $ExcludedObject -split " - *"
            $ExcludedUPN = $ExcludedUPN[($ExcludedUPN.count - 1)]
            [array]$PoliciesImpacted = $global:UserOutput | ? { ($_.user -eq $ExcludedUPN) -and ($_.Policy -eq $Policy.displayname) }

            foreach ($PolicyImpacted in $PoliciesImpacted) {

                $PolicyImpacted.'User Excluded' = "Yes"
                $PolicyImpacted.'Exclusion inherited from Group' = "Direct"

            }
        }
        ##Process policies excluding all guests or external users
        If ($ExcludedObject -eq "GuestsOrExternalUsers") {  
            foreach ($user in ($users | ? { $_.userprincipalname -like "*#EXT#@*" })) {
                [array]$PoliciesImpacted = $global:UserOutput | ? { ($_.user -eq $user.userPrincipalName) -and ($_.Policy -eq $Policy.displayname) }
                foreach ($PolicyImpacted in $PoliciesImpacted) {
                    $PolicyImpacted.'User Excluded' = "Yes"
                    if ($PolicyImpacted.'Exclusion inherited from Group' -eq "No") {
                        $PolicyImpacted.'Exclusion inherited from Group' = "All Guests or External Users"
                    }
                    else {
 
                        $PolicyImpacted.'Exclusion inherited from Group' = ($PolicyImpacted.'Exclusion inherited from Group' + " & All Guests or External Users")
                    }
                }
            }

        }

    }
    $ProgressStatus = "Processing policy: $($policy.displayName) - Group exclusions"
    UpdateProgress
    ##Assess Group excluded users
    foreach ($ExcludedGroups in ($policy.conditions.users.excludegroups)) {

        foreach ($Group in $ExcludedGroups) {
                
            $GroupSplit = $group -split " - *"
            $GroupID = $GroupSplit[($groupsplit.count - 1)]
            $apiuri = "https://graph.microsoft.com/v1.0/groups/$GroupID/members"
            $PassthroughError = "Error getting members for $GroupID - Group has possibly been deleted and should be removed from the policy $($policy.displayName)"
            [array]$Members = RunQueryandEnumerateResults
    
            foreach ($member in $Members) {
    
                if ($member.'@odata.type' -eq "#microsoft.graph.user") {
                    
                    $ExcludedUPN = $member.userprincipalname

                    [array]$PoliciesImpacted = $global:UserOutput | ? { ($_.user -eq $ExcludedUPN) -and ($_.Policy -eq $Policy.displayname) }
        
                    foreach ($PolicyImpacted in $PoliciesImpacted) {
                        
                        if ($PolicyImpacted.'Exclusion inherited from Group' -eq "No") {
                            $PolicyImpacted.'User Excluded' = "Yes"
                            $PolicyImpacted.'Exclusion inherited from Group' = $Group
                        }
                        else {
                            $PolicyImpacted.'Exclusion inherited from Group' = ($PolicyImpacted.'Exclusion inherited from Group' + " & $Group")
                        }
        
                    }

                }
                elseif ($member.'@odata.type' -eq "#microsoft.graph.group") {
                   
                    ##Expand nested groups
                    $global:grouplist += $Group
                    ExpandExclusionGroup -GroupDetails "$($member.displayName) - $($member.id)"
                    $global:grouplist = @()
                    $global:grouplistPrevious = @()
                }
            }
        }
    }##Update Progress
    $ProgressStatus = "Processing policy: $($policy.displayName) - Role exclusions"
    UpdateProgress
    ##Assess role exclusions
    foreach ($role in $policy.conditions.users.excluderoles) {
        $roleid = $null
        $roleid = ($EnabledDirectoryRoles | ? { $_.displayname -eq $role }).id
        if ($roleid) {
            $apiURI = "https://graph.microsoft.com/v1.0/directoryRoles/roleTemplateId=$roleid/members"
            [array]$Rolemembers = RunQueryandEnumerateResults

            foreach ($Rolemember in $Rolemembers) {
                $ExcludedUPN = $Rolemember.userPrincipalName
                [array]$PoliciesImpacted = $global:UserOutput | ? { ($_.user -eq $ExcludedUPN) -and ($_.Policy -eq $Policy.displayname) }

                foreach ($PolicyImpacted in $PoliciesImpacted) {
                        
                    if ($PolicyImpacted.'Exclusion inherited from Role Assignment' -eq "No") {
                        $PolicyImpacted.'User Excluded' = "Yes"
                        $PolicyImpacted.'Exclusion inherited from Role Assignment' = $Role
                    }
                    else {
                        $PolicyImpacted.'Exclusion inherited from Role Assignment' = ($PolicyImpacted.'Exclusion inherited from Role Assignment' + " & $Role")
                    }
    
                }

                $global:UserOutput += $row

            }

        }
    }
    
}

##Output to Excel file
$global:UserOutput | select 'User', 'User Excluded', 'Apps', 'Policy', 'Policy State', 'Inherited from Group', 'Exclusion inherited from Group', 'Inherited from Role Assignment', 'Exclusion inherited from Role Assignment'  | Export-Excel -Path ("$FilePath\$Filename") -WorksheetName "User Policies" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow


