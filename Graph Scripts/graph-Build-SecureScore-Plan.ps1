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

function CreateGroup {
    <#
    .SYNOPSIS
    Provisions Office 365 Group to contain Plan
    
    .PARAMETER token
    -token is the auth token
    
    #>
    Param(
        [parameter(Mandatory = $true)]
        $token

    )
    

    ##Build Group Request Body
    $RequestBody = @"
    {
        'description': 'SecureScore Tracking',
        'displayName': 'SecureScore Tracking',
        'groupTypes': [
          'Unified'
        ],
        'mailEnabled': true,
        'mailNickname': 'SecureScoreTracking',
        'securityEnabled': false,
        'members@odata.bind': [
            'https://graph.microsoft.com/v1.0/me'
          ]
    }
"@
    


    $apiUri = "https://graph.microsoft.com/beta/Groups"
    ##Invoke Group Request
    $Group = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)" } -ContentType 'application/json' -Body $RequestBody -Uri $apiUri -Method Post)

    return $Group
}


function CreatePlan {
    <#
    .SYNOPSIS
    Provisions a Plan in the created group. Returns the Plan object
    
    .PARAMETER token
    -token is the auth token
    
    .PARAMETER token
    -the Group ID of the Group created for the planner instance
    #>
    Param(
        [parameter(Mandatory = $true)]
        $token,
        [parameter(Mandatory = $true)]
        $GroupID

    )
    $RequestBody = @"

    {
        'owner': '$($group.id)',
        'title': 'SecureScore Tracking'
      }
"@
    


    $apiUri = "https://graph.microsoft.com/beta/planner/plans"
    ##Invoke Group Request
    $Plan = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)" } -ContentType 'application/json' -Body $RequestBody -Uri $apiUri -Method Post)


    return $Plan

}

function CreateBuckets {
    <#
    .SYNOPSIS
    Provisions Buckets for each control category. 
    
    .PARAMETER token
    -token is the auth token
    
    .PARAMETER categories
    -the Catagories to provision as buckets

    .PARAMETER plan
    -the plan to create the buckets in
    #>
    Param(
        [parameter(Mandatory = $true)]
        $token,
        [parameter(Mandatory = $true)]
        $categories,
        [parameter(Mandatory = $true)]
        $Plan

    )

    $Buckets = @()
    foreach ($category in $categories) {
        $RequestBody = @"

    {
        'name': '$category',
        'planId': '$($plan.id)',
        'orderHint': ' !'
      }
"@
        write-host $RequestBody
        $apiUri = "https://graph.microsoft.com/beta/planner/buckets"
        ##Invoke Group Request
        $Bucket = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)" } -ContentType 'application/json' -Body $RequestBody -Uri $apiUri -Method Post)
        $buckets += $bucket
    }

    return $Buckets
}

function CreateTasks {
    <#
    .SYNOPSIS
    Provisions tasks for each item in the relevent Buckets. 
    
    .PARAMETER token
    -token is the auth token
    
    .PARAMETER Bucket
    -the bucket to associate the task to

    .PARAMETER item
    -the task object
    #>
    Param(
        [parameter(Mandatory = $true)]
        $token,
        [parameter(Mandatory = $true)]
        $Item,
        [parameter(Mandatory = $true)]
        $Bucket

    )


    ##Build Task Request
    $RequestBody = @"
    {
        "planId": '$($Bucket.planid)',
        "bucketId": '$($Bucket.id)',
        "title": '$($Item.controlname.tostring())'
      }
"@


    $apiUri = "https://graph.microsoft.com/v1.0/planner/tasks"
    ##Create Task
    $task = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)" } -ContentType 'application/json' -Body $RequestBody -Uri $apiUri -Method Post) 

    ##Add task details
    $RequestBody = @"
        {
        "description": "$($Item.description)"
          }
"@

    $apiUri = "https://graph.microsoft.com/v1.0/planner/tasks/$($task.id)/details"
    $taskdetails = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.accesstoken)" } -Uri $ApiUri -Method Get)

    Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)";'If-Match'=$taskdetails.'@odata.etag'} -ContentType 'application/json' -Body $RequestBody -Uri $apiUri -Method Patch

}


function ConvertSecurescoretoPlan{
    <#
    .SYNOPSIS
    Provisions tasks for each item in the relevent Buckets. 
    
    .PARAMETER tenantID
    -The directory ID of your tenancy
    
    .PARAMETER clientID
    -the app reg client ID

    #>
    Param(
        [parameter(Mandatory = $true)]
        $tenantID,
        [parameter(Mandatory = $true)]
        $clientID


    )

##Get graph token using MSAL.PS
$token = GetDelegatedGraphToken -clientId $clientID -tenantID $tenantID 

##Create a Group to hold the plan
$Group = CreateGroup -token $token

##Wait one minute for Group Provisioning to finish
write-host "Waiting for Group to be provisioned"
start-sleep -Seconds 60

##Create Plan in new group
$Plan = CreatePlan -token $token -GroupID $group.id

##Get SecureScore List
$apiUri = "https://graph.microsoft.com/beta/security/secureScores"
$results = RunQueryandEnumerateResults -apiUri $apiuri -token $token.AccessToken 

##Create Buckets for each category
$Buckets = CreateBuckets -categories ($results[0].controlscores.controlcategory | select -Unique) -token $token -Plan $Plan

##Loop through each control entry
foreach ($Entry in $results[0].controlscores) {

    ##Match bucket to task
    $Bucket = $buckets | Where-Object { $_.name -eq $entry.controlCategory }
    ##create a task in the relevent bucket

    CreateTasks -token $token -Item $entry -Bucket $Bucket

}

}