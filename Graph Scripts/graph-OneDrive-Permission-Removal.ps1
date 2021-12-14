##Author: Sean McAvinue
##Details: Used as a Graph/PowerShell example, 
##          NOT FOR PRODUCTION USE! USE AT YOUR OWN RISK

<#
    .SYNOPSIS
    Main function, reports on file and folder structure in OneDrive for all imported users

    .PARAMETER User
    -UserDisplayName is the display name for the guest account

    .PARAMETER Delegate
    -UserEmail is the email address of the requested user
    
    .PARAMETER clientSecret
    -is the app registration client secret

    .PARAMETER clientID
    -is the app clientID

    .PARAMETER tenantID
    -is the directory ID of the tenancy
    
    .Example
    .\graph-OneDrive-Permission-Removal.ps1 -User username@contoso.com -Delegate username2@contoso.com -clientID $clientID -clientSecret $clientSecret -tenantID $tenantID
     
    #>
Param(
    [parameter(Mandatory = $true)]
    $clientId,
    [parameter(Mandatory = $true)]
    $tenantId,
    [parameter(Mandatory = $true)]
    $clientSecret,
    [parameter(Mandatory = $true)]
    $User,
    [parameter(Mandatory = $true)]
    $Delegate
)
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
    

function expandfolders {
    <#
    .SYNOPSIS
    Expands folder structure and sends files to be written and folders to be expanded
  
    .PARAMETER folder
    -Folder is the folder being passed
    
    .PARAMETER FilePath
    -filepath is the current tracked path to the file
    
    .NOTES
    General notes
    #>
    Param(
        [parameter(Mandatory = $true)]
        $folder,
        [parameter(Mandatory = $true)]
        $FilePath

    )



    write-host retrieved $filePath -ForegroundColor green
    $filepath = ($filepath + '/' + $folder.name)
    write-host $filePath -ForegroundColor yellow
    $apiUri = ('https://graph.microsoft.com/beta/users/' + $user + '/drive/root:' + $FilePath + ':/children')

    $Data = RunQueryandEnumerateResults -ApiUri $apiUri -Token $token

    ##Loop through Root folders
    foreach ($item in $data) {
        ##Remove permissions if present
        RemovePermissions -item $item
        ##IF Folder
        if ($item.folder) {
            
            write-host $item.name is a folder, passing $filePath as path
            expandfolders -folder $item -filepath $filepath

            
        }

    }


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

function RemovePermissions {
    <#
    .SYNOPSIS
    Removes Permissions where present for the delegate is present

    #>
    Param(
        [parameter(Mandatory = $true)]
        $item
    )
    $apiuri = "https://graph.microsoft.com/beta/users/$user/drive/items/$($item.id)/permissions"
    write-host "Checking permission from $($item.name) for user $user with ID $delegateID" -ForegroundColor Yellow
    ##Pass to run query function
    $Permissions = RunQueryandEnumerateResults -token $token -apiUri $apiuri

    foreach ($permission in $permissions) {
        if ($permission.grantedto.user.id -eq $DelegateID) {

            write-host "Removing permission from $($item.name) for user $user with ID $delegateID" -ForegroundColor green
            $apiUri = "https://graph.microsoft.com/v1.0/drives/$DriveID/items/$($Item.id)/permissions/$($permission.id)"
            write-host $apiuri
            Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Delete
        }
    }
}







    
#Generate Token
$token = GetGraphToken -clientID $clientId -TenantID $tenantId -clientSecret $clientSecret

##Get User ID for Delegate
$apiUri = 'https://graph.microsoft.com/v1.0/users/' + $Delegate
$DelegateID = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get).id

##Get Drive ID
$apiUri = 'https://graph.microsoft.com/v1.0/users/' + $User + '/drive/'
$DriveID = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get).id
##Query Site to get Site ID
$apiUri = 'https://graph.microsoft.com/v1.0/users/' + $User + '/drive/root/children'
$Data = RunQueryandEnumerateResults -ApiUri $apiUri -Token $token

##Loop through Root folders
ForEach ($item in $data) {

    $apiuri = "https://graph.microsoft.com/beta/users/$user/drive/items/$($item.ID)/permissions"
    ##Pass to run uery function
    $Permissions = RunQueryandEnumerateResults -token $token -apiUri $apiuri
        
    RemovePermissions -item $item
        
    ##IF Folder, then expand folder
    if ($item.folder) {
        $token = GetGraphToken -clientID $clientId -TenantID $tenantId -clientSecret $clientSecret
        write-host $item.name is a folder
        $filepath = ""
        expandfolders -folder $item -filepath $filepath

    }
 
}
    
    
    
