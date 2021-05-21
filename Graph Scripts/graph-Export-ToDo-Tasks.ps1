##Author: Sean McAvinue
##Details: Used as a Graph/PowerShell example, 
##          NOT FOR PRODUCTION USE! USE AT YOUR OWN RISK
##          Returns a list of ToDo Tasks for users imported from CSV
function GetDelegatedGraphToken {

    <#
    .SYNOPSIS
    Azure AD OAuth Application Token for Graph API
    Get OAuth token for a AAD Application using delegated permissions via the MSAL.PS library(returned as $token)

    .PARAMETER clientID
    -is the app clientID

    .PARAMETER tenantID
    -is the directory ID of the tenancy

    .PARAMETER redirectURI
    -is the redirectURI specified in the application registration, default value is https://localhost

    #>

    # Application (client) ID, tenant ID and secret
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $clientID,
        [parameter(Mandatory = $true)]
        [String]
        $tenantID,
        [parameter(Mandatory = $false)]
        $RedirectURI = "https://localhost"
    )
    import-module msal.ps
    $Token = Get-MsalToken -DeviceCode -ClientId $clientID -TenantId $tenantID -RedirectUri $RedirectURI

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

function exporttasks{
    <#
    .SYNOPSIS
    Lists tasks in a list and exports to CSV    
    .PARAMETER ListID
    -ListID is the current list

    .PARAMETER Token
    -token is the auth token

    .PARAMETER UPN
    -The UPN of the current user
    #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $apiUri,
        [parameter(Mandatory = $true)]
        $token,
        [parameter(Mandatory = $true)]
        $UPN
    )

    $Results = RunQueryandEnumerateResults -token $token.accesstoken -apiUri $apiUri
   write-host results: $results

}

function get-tasks {
    <#
    .SYNOPSIS
    Accepts a CSV with the heading 'userprincipalname' list of users UPNs and exports tasks for each user to C:\output - please ensure the folder exists
    
    .PARAMETER CSV
    -the path to the import CSV
    
    .PARAMETER clientID
    -is the app clientID

    .PARAMETER tenantID
    -is the directory ID of the tenancy

    .PARAMETER redirectURI
    -is the redirectURI specified in the application registration, default value is https://localhost

    #>

    # Application (client) ID, tenant ID and secret
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $CSV,
        [parameter(Mandatory = $true)]
        [String]
        $clientID,
        [parameter(Mandatory = $true)]
        [String]
        $tenantID,
        [parameter(Mandatory = $false)]
        $RedirectURI = "https://localhost"
    )

    ##Retrieve Delegated Access Token
    $token = GetDelegatedGraphToken -clientID $clientID -tenantID $tenantID -RedirectURI $RedirectURI
    $Userlist = import-csv $CSV
    foreach($user in $Userlist){
        ##API for User Query
        $apiUri = "https://graph.microsoft.com/v1.0/users/$($user.userprincipalname)"
        ##Run User Query
        $UserObject = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)" } -Uri $apiUri -Method Get)
        ##APIURI for Task Query
        $apiUri = "https://graph.microsoft.com/v1.0/users/$($UserObject.id)/todo/lists/Tasks/tasks"

           
        exporttasks -apiUri $apiuri -token $token -UPN $user.userprincipalname    

    
    }

}


