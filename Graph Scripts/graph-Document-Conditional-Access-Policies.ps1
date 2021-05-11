
##Author: Sean McAvinue
##Details: Exports conditional access policies to JSON 
##          USE AT YOUR OWN RISK
function GetGraphToken {

    <#
    .SYNOPSIS
    Azure AD OAuth Application Token for Graph API
    Get OAuth token for a AAD Application (returned as $token)
    
    #>

    # Application (client) ID, tenant ID and secret
    Param(
        [parameter(Mandatory = $true)]
        $clientId,
        [parameter(Mandatory = $true)]
        $tenantId,
        [parameter(Mandatory = $true)]
        $clientSecret

    )
    
    
    
    # Construct URI
    $uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
     
    # Construct Body
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }
     
    # Get OAuth 2.0 Token
    $tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
     
    # Access Token
    $token = ($tokenRequest.Content | ConvertFrom-Json).access_token
    
    #Returns token
    return $token
}
    



function RunQueryandEnumerateResults {
    <#
    .SYNOPSIS
    Runs Graph Query and if there are any additional pages, parses them and appends to a single variable
    
    .PARAMETER apiUri
    -APIURi is the apiUri to be passed
    
    .PARAMETER token
    -token is the auth token
    
    #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $apiUri,
        [parameter(Mandatory = $true)]
        $token

    )

    #Run Graph Query
    $Results = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)
    #Output Results for debug checking
    #write-host $results

    #Begin populating results
    $ResultsValue = $Results.value

    #If there is a next page, query the next page until there are no more pages and append results to existing set
    if ($results."@odata.nextLink" -ne $null) {
        write-host enumerating pages -ForegroundColor yellow
        $NextPageUri = $results."@odata.nextLink"
        ##While there is a next page, query it and loop, append results
        While ($NextPageUri -ne $null) {
            $NextPageRequest = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $NextPageURI -Method Get)
            $NxtPageData = $NextPageRequest.Value
            $NextPageUri = $NextPageRequest."@odata.nextLink"
            $ResultsValue = $ResultsValue + $NxtPageData
        }
    }

    ##Return completed results
    return $ResultsValue

    
}
function FindUser {
    <#
    .SYNOPSIS
    Translates Users from ObjectID to UPN
    
    #>
    # Access Token and user ObjectID
    Param(
        [parameter(Mandatory = $true)]
        $Token,
        [parameter(Mandatory = $true)]
        $User
    )
    $apiUri = "https://graph.microsoft.com/v1.0/users/$User"
    $UserObject = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)


    Return $UserObject.userprincipalname
}

function FindGroup {
    <#
    .SYNOPSIS
    Translates Groups from ObjectID to Group Name
    
    #>
    # Access Token and user ObjectID
    Param(
        [parameter(Mandatory = $true)]
        $Token,
        [parameter(Mandatory = $true)]
        $Group
    )
    $apiUri = "https://graph.microsoft.com/v1.0/Groups/$Group"
    $GroupObject = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)
    Return $GroupObject.displayname

}
function PerformTranslation {
    <#
    .SYNOPSIS
    Translates Users, Groups and Apps from Object IDs to Friendly Names
    
    #>
    # Access Token and Current Policy
    Param(
        [parameter(Mandatory = $true)]
        $Token,
        [parameter(Mandatory = $true)]
        $Policy
    )
    
    ##Collect Existing Service Principals
    $apiUri = "https://graph.microsoft.com/v1.0/serviceprincipals"
    $ServicePrincipals = RunQueryandEnumerateResults -token $token -apiUri $apiUri

    ##Process Included Applications
    foreach ($includeapp in $policy.conditions.applications.includeapplications) {

        if (($includeapp -ne "Office365") -and ($includeapp -ne "All")) {
            $includeapplist += (($ServicePrincipals | ? { $_.appid -eq $includeapp }).appDisplayName + ",")
        }
        elseif ($includeapp -eq "Office365") {
            $includeapplist += "Office365," 
        }
        elseif ($includeapp -eq "All") {
            $includeapplist += "All," 
        }

    }

    ##Process Excluded Applications
    foreach ($excludeapp in $policy.conditions.applications.excludeapplications) {

        if (($excludeapp -ne "Office365") -and ($includeapp -ne "All")) {
            $excludeapplist += (($ServicePrincipals | ? { $_.appid -eq $includeapp }).appDisplayName + ",")
        }
        elseif ($excludeapp -eq "Office365") {
            $excludeapplist += "Office365," 
        }
        elseif ($excludeapp -eq "All") {
            $excludeapplist += "All," 
        }
    }

    ##Process Included users
    foreach ($includeuser in $policy.conditions.users.includeusers) {
        if ($includeuser -eq "All") {
            $includeUserList = "All"

        }
        else {
            $UPN = FindUser -Token $token -User $includeuser
            $includeUserList += ( "$UPN," )

        }
    }

    ##Process Excluded users
    foreach ($excludeuser in $policy.conditions.users.excludeusers) {
        if ($excludeuser -eq "All") {
            $excludeUserList = "All"
        }
        else {
            $UPN = FindUser -Token $token -User $excludeuser
            $excludeUserList += ( "$UPN," )
                
        }
    }


    ##Process Included Groups
    foreach ($includegroup in $policy.conditions.users.includegroups) {
        if ($includegroup -eq "All") {
            $includegroupList = "All"

        }
        else {
            $GroupName = FindGroup -Token $token -group $includegroup
            $includegroupList += ( "$GroupName," )

        }
    }

        
    ##Process Excluded Groups
    foreach ($excludegroup in $policy.conditions.users.excludegroups) {
        if ($excludegroup -eq "All") {
            $excludeGroupList = "All"
        }
        else {
            $GroupName = FindGroup -Token $token -group $excludegroup

            $excludeGroupList += ( "$GroupName," )
                    
        }
    }





    $policy.conditions.applications.includeapplications = $includeapplist
    $policy.conditions.applications.excludeapplications = $excludeapplist
    $policy.conditions.users.includeusers = $includeuserlist
    $policy.conditions.users.excludeusers = $excludeuserlist
    $policy.conditions.users.includegroups = $includegrouplist
    $policy.conditions.users.excludegroups = $excludegrouplist

    return $policy
    

}
function Report-ConditionalAccess {

    <#
    .SYNOPSIS
    Returns a report of Conditional Access Policies in a tenent
    
    #>

    # Application (client) ID, tenant ID and secret
    Param(
        [parameter(Mandatory = $true)]
        $clientId,
        [parameter(Mandatory = $true)]
        $tenantId,
        [parameter(Mandatory = $true)]
        $clientSecret,
        [parameter(Mandatory = $false)]
        $PerformTranslation = $False

    )

    $apiUri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"
    $token = GetGraphToken -clientId $clientId -tenantId  $tenantId -clientSecret $clientSecret

    $Policies = RunQueryandEnumerateResults -apiuri $apiUri -token $token



    foreach ($policy in $policies) {


        if ($PerformTranslation) {

            $Policy = PerformTranslation -Token $token -Policy $Policy

        }

        $policy | convertto-json | out-file ("$($policy.displayName).json").replace('[', '').replace(']', '').replace('/', '')

    }

}


