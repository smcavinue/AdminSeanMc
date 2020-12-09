##Author: Sean McAvinue
##Details: Used as a Graph/PowerShell example, 
##          NOT FOR PRODUCTION USE! USE AT YOUR OWN RISK
##          Exports Planner instances to JSON files
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
    write-host running $apiuri -foregroundcolor blue

    $Results = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get -SkipHttpErrorCheck)

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

function ListGroups {
    <#
    .SYNOPSIS
    Runs Graph Query to list groups in the tenant
    
    .PARAMETER token
    -token is the auth token
    
    #>
    Param(
        [parameter(Mandatory = $true)]
        $token

    )
    ##Gets Unified Groups
    $apiUri = "https://graph.microsoft.com/beta/groups/?`$filter=groupTypes/any(c:c+eq+'Unified')"
    $Grouplist = RunQueryandEnumerateResults -token $token.accesstoken -apiUri $apiUri

    Write-host Found $grouplist.count Groups to process -foregroundcolor yellow

    Return $Grouplist

}


function SetGroupOwnership {
    <#
    .SYNOPSIS
    Runs Graph Query to list groups in the tenant
    
    .PARAMETER token
    -token is the auth token
    
    .PARAMETER GroupList
    -List of unified Groups in the tenant
    #>
    Param(
        [parameter(Mandatory = $true)]
        $token,
        [parameter(Mandatory = $true)]
        $Grouplist
    )

    foreach ($Group in $Grouplist) {

        $RequestBody = @"
        {

            "@odata.id": "https://graph.microsoft.com/v1.0/me"

        }
"@
        
    
        write-host Adding account as owner of $group.id
        $apiUri = "https://graph.microsoft.com/beta/Groups/$($Group.id)/owners/`$ref"
        ##Invoke Group Request
        $Group = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)" } -ContentType 'application/json' -Body $RequestBody -Uri $apiUri -Method Post)
    }
    foreach ($Group in $Grouplist) {

        $RequestBody = @"
        {

            "@odata.id": "https://graph.microsoft.com/v1.0/me"

        }
"@
        

        write-host Adding account as member of $group.id
        $apiUri = "https://graph.microsoft.com/beta/Groups/$($Group.id)/members/`$ref"
        ##Invoke Group Request
        $Group = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)" } -ContentType 'application/json' -Body $RequestBody -Uri $apiUri -Method Post)



    }

}

function exportplanner {
    <#
    .SYNOPSIS
    This function gets Graph Token from the GetGraphToken Function and uses it to request a new guest user

    .PARAMETER clientID
    -is the app clientID

    .PARAMETER tenantID
    -is the directory ID of the tenancy
    

    #>
    Param(
        [parameter(Mandatory = $true)]
        $clientId,
        [parameter(Mandatory = $true)]
        $tenantId
    )
    
    #Generate Token
    $token = GetDelegatedGraphToken -clientID $clientId -TenantID $tenantId

    $Grouplist = ListGroups -token $token

    SetGroupOwnership -token $token -grouplist $grouplist


    ##Loop through Groups in CSV
    foreach ($Group in $Grouplist) {

        ##Build Query
        $apiUri = "https://graph.microsoft.com/beta/groups/$($Group.id)/planner/plans"

        $Plans = RunQueryandEnumerateResults -apiUri $apiUri -token $token.accesstoken
        
        
        if ($plans) {

            $plans | Add-Member -Type NoteProperty -Name GroupID -Value $Group.id

            $plans |  export-csv planslist.csv -NoClobber -NoTypeInformation -Append

            foreach ($Plan in $plans) {

                $apiUri = "https://graph.microsoft.com/beta/planner/plans/$($plan.id)/buckets"
                $buckets = RunQueryandEnumerateResults -apiUri $apiUri -token $token.accesstoken
                if ($buckets) {
                    $buckets | export-csv "$($plan.id)-buckets.csv" -NoClobber -NoTypeInformation -Append
                }
            
            }

            foreach ($Plan in $plans) {

                $apiUri = "https://graph.microsoft.com/beta/planner/plans/$($plan.id)/tasks"
    
                $tasks = RunQueryandEnumerateResults -apiUri $apiUri -token $token.accesstoken
                if ($tasks) {
                    $tasks  | export-csv "$($plan.id)-tasks.csv" -NoClobber -NoTypeInformation -Append
                }
                
            }
        }
    }
    


}