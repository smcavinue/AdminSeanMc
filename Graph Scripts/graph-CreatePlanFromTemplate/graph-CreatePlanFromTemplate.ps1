##Parameters for the script to accept clientID, tenantID, clientSecret, CSVfilepath, PlanName and GroupID
param(
    [parameter(Mandatory = $true)]
    [String]
    $clientID,
    [parameter(Mandatory = $true)]
    [String]
    $tenantID,
    [parameter(Mandatory = $true)]
    [String]
    $clientSecret,
    [parameter(Mandatory = $true)]
    [String]
    $csvFilePath,
    [parameter(Mandatory = $true)]
    [String]
    $PlanName,
    [parameter(Mandatory = $true)]
    [String]
    $GroupID,
    [parameter(Mandatory = $false)]
    [String]
    $StorageAccountName,
    [parameter(Mandatory = $false)]
    [String]
    $StorageContainerName,
    [parameter(Mandatory = $false)]
    [String]
    $TeamsChannelName
)

##Example for running locally:
##.\graph-CreatePlanFromTemplate.ps1 -clientID $clientID -tenantID $tenantID -clientSecret $clientSecret -csvfilepath $csvfilepath -PlanName "Tenant to Tenant Migration Plan" -GroupID $groupID
##StorageAccountName and StorageContainerName are optional parameters for running from Azure Automation
If ($StorageContainerName) {

    write-host "Storage Container Parameter detected, downloading CSV from storage account"
    Connect-AZAccount -Identity

    $context = New-AzStorageContext -StorageAccountName $storageaccountname
   
    Get-AzStorageBlobContent -Blob $CSVFilePath -Container $StorageContainerName -Context $context
   
    $csv = import-csv $csvFilePath

}
else {
    ##Import the CSV
    Try {
        $csv = import-csv $csvFilePath
    }
    catch {
        write-error "Could not import CSV, please check the path and try again. Error:`n $_"
        exit
    }
}
##Connect to Microsoft Graph

try {
    $body = @{
        Grant_Type    = "client_credentials"
        Scope         = "https://graph.microsoft.com/.default"
        Client_Id     = $clientID
        Client_Secret = $clientSecret
    }
 
    $TokenRequest = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token" -Method POST -Body $body
 
    $token = $TokenRequest.access_token
 
    Connect-MgGraph -AccessToken $token

}
catch {
    write-error "Could not obtain a Microsoft Graph token. Error:`n $_"
    exit
}

write-host "Provisioning Plan"
##Create the Plan
$params = @{
    container = @{
        url = "https://graph.microsoft.com/v1.0/groups/$groupid"
    }
    title     = $planName
}
        
try {
    $plan = New-MgPlannerPlan -BodyParameter $params
}
catch {
    write-error "Could not create the plan. Error:`n $_"
    exit
}

write-host "Provisioning Buckets"
##Loop through the unique buckets and provision the Buckets
[array]$Buckets = ($csv | Select-Object Bucket -Unique)
$orderhint = " !"
$BucketList = @()
$i = 0
foreach ($bucket in $Buckets) {
    $i++
    Write-Progress -Activity "Creating Buckets" -Status "Creating Bucket $i of $($Buckets.count)" -PercentComplete (($i / $Buckets.count) * 100)
    $params = @{
        name      = "$($bucket.Bucket)"
        planId    = "$($plan.id)"
        orderHint = "$($orderhint)"
    }
    
    $CreatedBucket = New-MgPlannerBucket -BodyParameter $params
    $BucketList += $CreatedBucket
    $orderhint = " $($createdBucket.orderhint)!"

}

write-host "Provisioning Tasks"
##Create Tasks in buckets
$i = 0
foreach ($Task in $csv) {
    $i++
    Write-Progress -Activity "Creating Tasks" -Status "Creating Task $i of $($csv.count)" -PercentComplete (($i / $csv.count) * 100)
    $CurrentBucket = $BucketList | Where-Object { $_.name -eq $Task.Bucket }

    try {
        
        $params = @{
            planId   = "$($Plan.id)"
            bucketId = "$($CurrentBucket.id)"
            title    = "$($Task.task)"
        }
        
        $CreatedTask = New-MgPlannerTask -BodyParameter $params
    }
    catch {
        write-error "Could not create task: $($task.task), Error:`n $_"
        exit
    }

    $params = @{
        description = "$($Task.details)"
        previewType = "description"
    }
    ##Update Plan Details
    try {
        
        Update-MgPlannerTaskDetail -PlannerTaskId $CreatedTask.Id -BodyParameter $params -IfMatch (Get-MgPlannerTaskDetail -PlannerTaskId $CreatedTask.id).AdditionalProperties["@odata.etag"] 
    }
    catch {
        write-error "Could not update task details: $($task.task), Error:`n $_"
        exit
    }
}

##Add Planner to Teams Channel 
if ($TeamsChannelName) {
    Try {
        $ChannelID = (Get-MgTeamChannel -TeamId $groupid | ?{$_.DisplayName -eq $TeamsChannelName}).id
        $params = @{
            name                  = $PlanName
            displayName           = $PlanName
            "teamsapp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.planner"
            configuration         = @{
                contentUrl = "https://tasks.teams.microsoft.com/teamsui/{tid}/Home/PlannerFrame?page=7&auth_pvr=OrgId&auth_upn={userPrincipalName}&groupId={groupId}&planId=$($plan.id)&channelId={channelId}&entityId={entityId}&tid={tid}&userObjectId={userObjectId}&subEntityId={subEntityId}&sessionId={sessionId}&theme={theme}&mkt={locale}&ringId={ringId}&PlannerRouteHint={tid}"
                removeUrl = "https://tasks.teams.microsoft.com/teamsui/{tid}/Home/PlannerFrame?page=13&auth_pvr=OrgId&auth_upn={userPrincipalName}&groupId={groupId}&planId=$($plan.id)&channelId={channelId}&entityId={entityId}&tid={tid}&userObjectId={userObjectId}&subEntityId={subEntityId}&sessionId={sessionId}&theme={theme}&mkt={locale}&ringId={ringId}&PlannerRouteHint={tid}"
                websiteUrl = "https://tasks.office.com/{tid}/Home/PlanViews/$($Plan.id)?Type=PlanLink&Channel=TeamsTab"        
            }
        
        }    
        $CreatedTab = New-MgTeamChannelTab -TeamId $groupid -ChannelId $ChannelID -BodyParameter $params
    }
    catch {
        write-error "Could not create tab for task: $($task.task), Error:`n $_"
        exit
    }
}