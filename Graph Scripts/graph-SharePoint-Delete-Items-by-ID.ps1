##Author: Sean McAvinue
##Details: Used as a Graph/PowerShell example, 
##          NOT FOR PRODUCTION USE! USE AT YOUR OWN RISK
##          Returns a report of OneDrive file and folder structure to CSV file
function GetGraphToken {
    <#
    .SYNOPSIS
    Azure AD OAuth Application Token for Graph API
    Get OAuth token for a AAD Application (returned as $token)
    
    #>

    # Application (client) ID, tenant ID and secret
    $clientId = "3b04d58f-9abe-4917-a129-033125d573ab"
    $tenantId = "845ee061-0700-412a-90f5-1465fb9a39f1"
    $clientSecret = "PdDf~OXb609K.s7v.l_thRS5~mhcKY7vam"
    
    
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
    $apiUri = ('https://graph.microsoft.com/beta/drives/' + $drive.id + '/root:' + $FilePath + ':/children')

    $Data = RunQueryandEnumerateResults -ApiUri $apiUri -Token $token

    ##Loop through Root folders
    foreach ($item in $data) {

        ##IF Folder
        if ($item.folder) {

            write-host $item.name is a folder, passing $filePath as path
            expandfolders -folder $item -filepath $filepath

            
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

    ##Build file object
    $object = [PSCustomObject]@{
        FileName     = $File.name
        LastModified = $File.lastModifiedDateTime
        Filepath     = $filepath
        Site         = $group.id
        ItemID      = $file.id
    }

    ##Export File Object
    $object | export-csv SharePointReport.csv -NoClobber -NoTypeInformation -Append

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
    write-host $apiuri
    #Run Graph Query
    $Results = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)
    #Output Results for debug checking
    

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



    #>

    ##Get in scope Users from CSV file##
    $groups = import-csv .\CK_Delete.csv
    $i = 0

    #Loop Through Users
    foreach ($Group in $Groups) {
    
        $i++

        if($i -gt 500){
        #Generate Token
        $token = GetGraphToken
        $i=0
        }
        $apiuri = "https://graph.microsoft.com/beta/sites/mtuireland.sharepoint.com:/sites/$($group.id)"
        $site  = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)
        ##Query Site to get Site ID
        $apiUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives"
        #$apiUri = "https://graph.microsoft.com/v1.0/groups/f765602e-5e68-46b5-8b80-c6e22f51d448/drives"
        $Drive = RunQueryandEnumerateResults -ApiUri $apiUri -Token $token | ?{$_.name -like "Documents"}

        $apiuri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives/$($drive.id)/items/$($Group.itemid)"
        $DriveItem = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)

    
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $token" } -Uri $apiUri -Method Delete

    
    }
    

    
