##Author: Sean McAvinue
##Details: Used as a Graph/PowerShell example, 
##          NOT FOR PRODUCTION USE! USE AT YOUR OWN RISK
##          Exports Planner instances to CSV files
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
        $Group = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)" } -ContentType 'application/json' -Body $RequestBody -Uri $apiUri -Method Post -ErrorAction silentlycontinue) 
    }
    start-sleep 20
    foreach ($Group in $Grouplist) {

        $RequestBody = @"
        {

            "@odata.id": "https://graph.microsoft.com/v1.0/me"

        }
"@
        

        write-host Adding account as member of $group.id
        $apiUri = "https://graph.microsoft.com/beta/Groups/$($Group.id)/members/`$ref"
        ##Invoke Group Request
        $Group = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)" } -ContentType 'application/json' -Body $RequestBody -Uri $apiUri -Method Post  -ErrorAction silentlycontinue)



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
        $tenantId,
        [parameter(Mandatory = $false)]
        $Grouplistfile
    )
    
    #Generate Token
    $token = GetDelegatedGraphToken -clientID $clientId -TenantID $tenantId

    IF(!$grouplistfile){
    $Grouplist = ListGroups -token $token
    }else{
    $Grouplist = import-csv $Grouplistfile
    }

    $grouplist
    #SetGroupOwnership -token $token -grouplist $grouplist


    ##Loop through Groups in CSV
    foreach ($Group in $Grouplist) {

        ##Build Query
        $apiUri = "https://graph.microsoft.com/beta/groups/$($Group.id)/planner/plans"

        $Plans = RunQueryandEnumerateResults -apiUri $apiUri -token $token.accesstoken
        
        
        if ($plans) {

            $plans | Add-Member -Type NoteProperty -Name GroupID -Value $Group.id

            $plans | export-csv planslist.csv -NoClobber -Append -NoTypeInformation

            foreach ($Plan in $plans) {
                $apiUri = "https://graph.microsoft.com/beta/planner/plans/$($plan.id)/details"
                $PlanDetails = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)" } -Uri $apiUri -Method Get)
                $PlanDetailsExport = [PSCustomObject]@{
                    categoryDescriptions = $PlanDetails.categoryDescriptions
                }

                
                $PlanDetailsExport  | ConvertTo-Json |  out-file "$($plan.id)-planDetails.json" -NoClobber -Append

                $apiUri = "https://graph.microsoft.com/beta/planner/plans/$($plan.id)/buckets"
                $buckets = RunQueryandEnumerateResults -apiUri $apiUri -token $token.accesstoken
                if ($buckets) {
                    $buckets | ConvertTo-Json |  out-file "$($plan.id)-buckets.json" -NoClobber -Append
                }
            
            }

            foreach ($Plan in $plans) {

                $apiUri = "https://graph.microsoft.com/beta/planner/plans/$($plan.id)/tasks"
    
                $tasks = RunQueryandEnumerateResults -apiUri $apiUri -token $token.accesstoken
                if ($tasks) {
                    $tasks  | ConvertTo-Json |  out-file "$($plan.id)-tasks.json" -NoClobber -Append

                    foreach($task in $tasks){

                        $apiUri = "https://graph.microsoft.com/beta/planner/tasks/$($task.id)/details"
                        $taskdetails = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)" } -Uri $apiUri -Method Get)
                        $taskdetails  | ConvertTo-Json |  out-file "$($task.id)-taskdetails.json" -NoClobber -Append
                        start-sleep 1
                    }
                }
                
            }
        }
    start-sleep 10
    }
    


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
        $GroupID,
        [parameter(Mandatory = $true)]
        $Title,
        [parameter(Mandatory = $true)]
        $Categories
    )
    $RequestBody = @"

    {
        'owner': '$($groupid)',
        'title': '$($title)'
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
    
    .PARAMETER buckets
    -the buckets to provision

    .PARAMETER plan
    -the plan to create the buckets in
    #>
    Param(
        [parameter(Mandatory = $true)]
        $token,
        [parameter(Mandatory = $true)]
        $buckets,
        [parameter(Mandatory = $true)]
        $Plan

    )

    $NewBuckets = @()
    foreach ($bucket in $buckets) {
        $RequestBody = @"

    {
        'name': '$($bucket.name)',
        'planId': '$($plan.id)',
        'orderHint': '$(" !")'
      }
"@
        write-host $RequestBody -Verbose
        $apiUri = "https://graph.microsoft.com/beta/planner/buckets"
        ##Invoke Group Request
        $Bucket = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)" } -ContentType 'application/json' -Body $RequestBody -Uri $apiUri -Method Post)
        $Newbuckets += $bucket
    }

    return $NewBuckets
}

function CreateTasks {
    <#
    .SYNOPSIS
    Provisions tasks for each item in the relevent Buckets. 
    
    .PARAMETER token
    -token is the auth token

    .PARAMETER taskBody
    -the task object
        
    .PARAMETER taskDetailsBody
    -the task details object
    #>
    Param(
        [parameter(Mandatory = $true)]
        $token,
        [parameter(Mandatory = $true)]
        $TaskBody,
        [parameter(Mandatory = $true)]
        $TaskDetailsBody
    )


 
    $apiUri = "https://graph.microsoft.com/v1.0/planner/tasks"
    ##Create Task
    write-host "Provisioning Task $taskbody END" -ForegroundColor green
    $TaskBody
    $task = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)" } -ContentType 'application/json' -Body $TaskBody -Uri $apiUri -Method Post) 

    start-sleep 5

    $apiUri = "https://graph.microsoft.com/v1.0/planner/tasks/$($task.id)/details"
    write-host "getting created task Details"
    $taskdetails = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.accesstoken)" } -Uri $ApiUri -Method Get)
    write-host "Provisioning Task Details $taskdetailsbody" -ForegroundColor yellow
    write-host $apiUri -ForegroundColor blue
    #pause
    Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token.AccessToken)";'If-Match'=$taskdetails.'@odata.etag'} -ContentType 'application/json' -Body $TaskDetailsBody -Uri $apiUri -Method Patch

}



function importplanner {
    <#
    .SYNOPSIS
    This function gets Graph Token from the GetGraphToken Function and uses it to request a new guest user

    .PARAMETER clientID
    -is the app clientID

    .PARAMETER tenantID
    -is the directory ID of the tenancy
    
    .PARAMETER GroupID
    -is the directory ID of the tenancy
    #>
    Param(
        [parameter(Mandatory = $true)]
        $clientId,
        [parameter(Mandatory = $true)]
        $tenantId,
        [parameter(Mandatory = $true)]
        $SourceGroupID,
        [parameter(Mandatory = $true)]
        $TargetGroupID,
        [parameter(Mandatory = $true)]
        $PlanID
    )
    
    #Generate Token
    $token = GetDelegatedGraphToken -clientID $clientId -TenantID $tenantId

    $csv = import-csv .\planslist.csv

    $PlanEntry = $csv | ?{(($_.groupid -eq $SourceGroupID) -and ($_.id -eq $planID))}
    $Bucketfile = Get-Content "$($PlanEntry.id)-buckets.json"
    $Taskfile = Get-Content "$($PlanEntry.id)-tasks.json"

    $Taskfile
    $Bucketfile

    $NewPlan = CreatePlan -token $token -GroupID $TargetGroupID -Title $PlanEntry.title

    
    $buckets = $Bucketfile | ConvertFrom-Json 

    $NewBuckets = CreateBuckets -token $token -buckets $buckets -Plan $NewPlan


    $tasks = $Taskfile | ConvertFrom-Json


    foreach($task in $tasks){

        $OldTaskBucket = $buckets | ?{$_.id -like $task.bucketid}

        $NewTaskBucket = $NewBuckets | ?{$_.name -like $OldTaskBucket.name}

        $TaskBody = @"
        {
            "planId": '$($NewTaskBucket.planid)',
            "bucketId": '$($NewTaskBucket.id)',
            "title": '$($task.title)',
            "percentComplete": '$($task.percentComplete)'
          }
"@


        $TaskDetailsFile = Get-Content "$($Task.id)-taskdetails.json"
        $TaskDetails = $TaskDetailsFile | ConvertFrom-Json 

        $checklists = Get-Member -InputObject $taskdetails.checklist | ?{$_.membertype -like "NoteProperty"}
        foreach($checklist in $checklists){
            $taskdetails.checklist.($checklist.name).orderHint = " !"
            $taskdetails.checklist.($checklist.name).lastModifiedBy = ""
            $taskdetails.checklist.($checklist.name).lastModifiedDateTime = ""
        }
        $TaskDetails.id = ""
        $TaskDetails.'@odata.context' = ""
        $TaskDetails.'@odata.etag' = ""
        $TaskDetails.references = ""

        $TaskDetailsBody = $TaskDetails | ConvertTo-Json
        $TaskDetailsBody = $taskdetailsbody.replace('"lastModifiedDateTime":  "",','')
        $TaskDetailsBody = $taskdetailsbody.replace('"lastModifiedby":  "",','')
        $TaskDetailsBody = $taskdetailsbody.replace('"lastModifiedBy":  ""','')
        $TaskDetailsBody = $taskdetailsbody.replace('"lastModifiedby":  "@{user=}"','')
        $TaskDetailsBody = $taskdetailsbody.replace('"references":  "",','')
        $TaskDetailsBody = $taskdetailsbody.replace('"orderHint":  " !",','"orderHint":  " !"')

        $newTasks = CreateTasks -token $token -TaskBody $TaskBody -TaskDetailsBody $TaskDetailsBody
    }
    



}