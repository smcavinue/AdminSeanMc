##Author: Sean McAvinue
##Details: Used as a Graph/PowerShell example, 
##          NOT FOR PRODUCTION USE! USE AT YOUR OWN RISK
##          Returns a report of OneDrive file and folder structure along with any sharing permissions to CSV file
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
    $apiUri = ('https://graph.microsoft.com/beta/users/' + $user.UserPrincipalName + '/drive/root:' + $FilePath + ':/children')

    $Data = RunQueryandEnumerateResults -ApiUri $apiUri -Token $token

    ##Loop through Root folders
    foreach ($item in $data) {

        ##IF Folder
        if ($item.folder) {

            write-host $item.name is a folder, passing $filePath as path
            expandfolders   -folder $item -filepath $filepath

            
        }##ELSE NOT Folder
        else {

            write-host $item.name is a file
            writeTofile -file $item -filepath $filePath

        }

    }


}
   
function writeTofile {
    <#
    .SYNOPSIS
    Writes files and paths to export file

    
    .PARAMETER File
    -file is the file name found
    
    .PARAMETER FilePath
    -filepath is the current tracked path
    
    #>
    Param(
        [parameter(Mandatory = $true)]
        $File,
        [parameter(Mandatory = $true)]
        $FilePath

    )

    ##If Shared, get the permissions
    if ($item.shared) {

        $Permissions = GetSharedFilePermissions -itemID $item.id -Token $token -itemname $item.name -Username $user.userprincipalname
                    
        write-host "found $($permissions.roles) permissions for $($file.name)" -ForegroundColor blue               
    }##Else blank out permission variable
    else {
        $permissions = $null
    }
    ##If there are multiple, build multiple objects and export each
    if ($Permissions -is [array]) {
        foreach ($permission in $permissions) {
            ##Build file object
            $object = [PSCustomObject]@{
                User              = $user.userprincipalname
                ID                = $item.id
                FileName          = $File.name
                shared            = $File.shared
                LastModified      = $File.lastModifiedDateTime
                Filepath          = $filepath
                ItemID            = $permission.itemID
                ItemName          = $permission.itemName
                hasPassword       = $permission.haspassword
                roles             = $permission.roles
                DirectPermissions = $permission.DirectPermissions
                LinkPermissions   = $permission.LinkPermissions
            }

            ##Export File Object
            $datestamp = (get-date).tostring('yyMMdd')
            $object | export-csv "OneDriveSharingReport-$datestamp.csv" -NoClobber -NoTypeInformation -Append


        }
    }
    else {
        ##Build file object
        $object = [PSCustomObject]@{
            User              = $user.userprincipalname
            ID                = $item.id
            FileName          = $File.name
            shared            = $File.shared
            LastModified      = $File.lastModifiedDateTime
            Filepath          = $filepath
            ItemID            = $permissions.itemID
            ItemName          = $permissions.itemName
            hasPassword       = $permissions.haspassword
            roles             = $permissions.roles
            DirectPermissions = $permissions.DirectPermissions
            LinkPermissions   = $permissions.LinkPermissions
        }

        ##Export File Object
        $datestamp = (get-date).tostring('yyMMdd')
        $object | export-csv "OneDriveSharingReport-$datestamp.csv" -NoClobber -NoTypeInformation -Append
    }
    ##Reset workingfilepath



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

function GetSharedFilePermissions {
    <#
    .SYNOPSIS
    Returns sharing details for input item
    
    .PARAMETER itemID
    -APIURi is the ID of the current item
    
    .PARAMETER itemName
    -token is the item to be processed
    
    .PARAMETER token
    -token is the auth token

    .PARAMETER username
    -token is current processed user
    #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $itemID,
        [parameter(Mandatory = $true)]
        [String]
        $itemName,
        [parameter(Mandatory = $true)]
        [String]
        $Token,
        [parameter(Mandatory = $true)]
        [String]
        $Username
    )
    ##Build Query
    $apiuri = "https://graph.microsoft.com/beta/users/$username/drive/items/$itemID/permissions"
    ##Pass to run uery function
    $Permissions = RunQueryandEnumerateResults -token $token -apiUri $apiuri
    ##Build an array to hold results
    $Permismissionarray = @()
    ##Loop through Permissions and create object to hold results. If there are multiple these will be appended to the array
    foreach ($permission in $permissions) {
        $PermissionObject = New-Object PSObject -Property @{
            ItemID            = $itemID
            ItemName          = $itemName
            hasPassword       = $permission.haspassword
            roles             = $permission.roles[0]
            DirectPermissions = $permission.grantedto.user.email -join (' ')
            LinkPermissions   = $permission.grantedtoidentities.user.email -join (' ')
        }
        $Permismissionarray += $PermissionObject
        
    }

    return $Permismissionarray 

}


function getonedrivereport {
    <#
    .SYNOPSIS
    Main function, reports on file and folder structure in OneDrive for all imported users

    #>
    Param(
        [parameter(Mandatory = $true)]
        $clientId,
        [parameter(Mandatory = $true)]
        $tenantId,
        [parameter(Mandatory = $true)]
        $clientSecret,
        [parameter(Mandatory = $true)]
        $UserListCSV
    )

    ##Get in scope Users from CSV file##
    $Users = import-csv $UserListCSV


    #Loop Through Users
    foreach ($User in $Users) {
    
        #Generate Token
        $token = GetGraphToken -clientID $clientId -TenantID $tenantId -clientSecret $clientSecret

        ##Query Site to get Site ID
        $apiUri = 'https://graph.microsoft.com/v1.0/users/' + $User.userprincipalname + '/drive/root/children'
        $Data = RunQueryandEnumerateResults -ApiUri $apiUri -Token $token

        ##Loop through Root folders
        ForEach ($item in $data) {

            ##IF Folder, then expand folder
            if ($item.folder) {

            
                write-host $item.name is a folder
                $filepath = ""
                expandfolders -folder $item -filepath $filepath

                ##ELSE NOT Folder, then it's a file, sent to write output
            }
            else {

                write-host $item.name is a file
                $filepath = ""
                writeTofile -file $item -filepath $filepath

            }

        }


    
    }
    
    
    
}