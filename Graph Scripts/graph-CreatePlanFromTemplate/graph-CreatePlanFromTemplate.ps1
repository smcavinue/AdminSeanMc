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
    $StorageContainerName
)


If ($StorageContainerName) {

    write-host "Storage Container Parameter detected, downloading CSV from storage account"
    Connect-AZAccount -Identity

    $context = New-AzStorageContext -StorageAccountName $storageaccountname
   
    Get-AzStorageBlobContent -Blob $CSVFilePath -Container $StorageContainerName -Context $context
   
    $csv = import-csv plantemplate.csv

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
    write-host $orderhint

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



