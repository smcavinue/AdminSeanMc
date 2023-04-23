##Author: Sean McAvinue
##Details: Used as a Graph/PowerShell example, 
##          NOT FOR PRODUCTION USE! USE AT YOUR OWN RISK
##          Returns a report of OneDrive file and folder structure to CSV file


Param(
    [parameter(Mandatory = $true)]
    [String]
    $SiteURL,
    [parameter(Mandatory = $true)]
    [String]
    $tenantID,
    [parameter(Mandatory = $true)]
    [String]
    $ClientID,
    [parameter(Mandatory = $true)]
    [String]
    $Secret
)

function GetGraphToken {
    <#
    .SYNOPSIS
    Azure AD OAuth Application Token for Graph API
    Get OAuth token for a AAD Application (returned as $token)
    
    #>

    
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
  
    .NOTES
    General notes
    #>


    $filepath = ($filepath + '/' + $item.name)


    $apiUri = ("https://graph.microsoft.com/beta/drives/$($Drive.id)/root:$($FilePath):/children")
    $items = RunQueryandEnumerateResults -apiUri $apiUri -token $token

    foreach ($item in $items) {
        CheckPermissions
    }

    $splitpath = $filepath.split('/')

    $filepath = $splitpath[0..($splitpath.count - 2)] -join ('/')
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
    Do {
        Try {
            $Results = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)
            $HeaderCheck = $null
        }
        catch {
            $HeaderCheck = $null
            $HeaderCheck = $error[0]
            if ($HeaderCheck.Exception.Response.StatusCode -eq 429) {
                write-host "Throttling encountered accessing $Apiuri, backing off temporarily" -ForegroundColor Yellow
                start-sleep -Seconds 30
                write-host "Resuming.." -ForegroundColor Yellow

            }
        }
    }Until($HeaderCheck.Exception.Response.StatusCode -ne 429)
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
            Do {
                Try {
                    $HeaderCheck = $null
                    $NextPageRequest = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $NextPageURI -Method Get)
                }
                catch {
                    $HeaderCheck = $null
                    $HeaderCheck = $error[0]
                    if ($HeaderCheck.Exception.Response.StatusCode -eq 429) {
                        write-host "Throttling encountered accessing next page for $Apiuri, backing off temporarily" -ForegroundColor Yellow
                        start-sleep -Seconds 30
                        write-host "Resuming.." -ForegroundColor Yellow
                    }
                }
                $NxtPageData = $NextPageRequest.Value
                $NextPageUri = $NextPageRequest."@odata.nextLink"
                $ResultsValue = $ResultsValue + $NxtPageData
            }Until($HeaderCheck.Exception.Response.StatusCode -ne 429)
        }
    }
    ##Return completed results
    return $ResultsValue
}

function CheckPermissions {
    <#
   .SYNOPSIS
   Fetches Permissions on an item and returns the inheritence status
    
    .PARAMETER token
    -token is the auth token

    .PARAMETER item
    -The driveitem object

    .PARAMETER item
    -The drive object

   #>


    if ([bool]($item.PSobject.Properties.name -match "folder")) {
        #Generate Token
        $token = GetGraphToken
        Write-Host "$($item.name) is a folder"

        ExpandFolders 
        
    }
    else {
        Write-Host "$($item.name) is a file"
    }

    $apiUri = "https://graph.microsoft.com/v1.0/drives/$($Drive.id)/items/$($item.id)/permissions"
    $Permissions = RunQueryandEnumerateResults -apiUri $apiUri -token $token
    $TestForInheritence = [bool]($Permissions[0].PSobject.Properties.name -match "inheritedFrom")
    if ($TestForInheritence) {
        write-host "Permissions inherited for $($item.name)"
    }
    else {
        write-host "Inheritance broken for $($item.name)" -ForegroundColor yellow
        $Export = @{
            "Site"     = "$SiteURL"
            "Filepath" = "$($Filepath)/$($item.name)"
            "Library" = "$($drive.name)"
        }
        [PSCustomObject]$Export | Export-CSV .\inheritanceReport-it.csv -NoClobber -NoTypeInformation -Append
    }
}


#Generate Token
$token = GetGraphToken   


Try {
    ##Query Site to get Site ID
    $siteSplit = $SiteURL.split('/')
    if ($SiteURL -like "*/sites/*") {

        $apiUri = "https://graph.microsoft.com/v1.0/sites/$($siteSplit[2]):/$($siteSplit[3])/$($siteSplit[4])/$($siteSplit[5])"
        
    }
    else {
        $apiUri = "https://graph.microsoft.com/v1.0/sites/$($siteSplit[2]):/$($siteSplit[3])"
    }
    $Site = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)

    $apiUri = "https://graph.microsoft.com/v1.0/sites/$($Site.id)/drives"
    $Drives = RunQueryandEnumerateResults -token $token -apiUri $apiUri
    $drive = $drives | Out-GridView -PassThru -Title "Select Drive to Process"
}
catch {
    write-error "Cannot retrieve site or drives for $siteURL"
    stop
}

$apiUri = "https://graph.microsoft.com/v1.0/drives/$($Drive.id)/root/children"
$items = RunQueryandEnumerateResults -apiUri $apiUri -token $token
$Tracker = 1

foreach ($item in $items) {
    Write-Progress -Activity "Scanning for inheritence" -Status "Processing Item $($item.name): $tracker / $($items.count)" -PercentComplete (($tracker / ($items.count)) * 100)
    $tracker++

    #Generate Token
    $token = GetGraphToken
    $filepath = ""
    CheckPermissions

}