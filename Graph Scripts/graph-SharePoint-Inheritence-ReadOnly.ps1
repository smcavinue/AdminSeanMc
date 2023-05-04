##Author: Sean McAvinue
##Details: Used as a Graph/PowerShell example, 
##          NOT FOR PRODUCTION USE! USE AT YOUR OWN RISK
##          Returns a report of OneDrive file and folder structure to CSV file


Param(
    [parameter(Mandatory = $true)]
    [String]
    $InputfilePath,
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
    




Try {
    $CSV = import-csv $InputfilePath
}
catch {
    Write-Host "Cannot find input file... exiting"
    pause
    exit
}


#Generate Token
$token = $null
$token = GetGraphToken   

If (!$token) {
    Write-Host "Cannot obtain toekn... exiting"
    pause
    exit
}

$drive = $null
$site = $null
$tracker = 1
foreach ($entry in $csv) {

    $token = GetGraphToken   

    Write-Progress -Activity "Making files and folders ready only in site $($entry.site)" -Status "Processing Item $($entry.filepath): $tracker / $($csv.count)" -PercentComplete (($tracker / ($csv.count)) * 100)

    $tracker++

    $SiteURL = $entry.site
    $LibraryName = $entry.library
    $Filepath = $entry.Filepath

    if (($drive.name -ne $LibraryName) -or ($Site.webUrl -ne $SiteURL)) {

    
        $SiteSplit = $siteurl.split('/')

        if ($SiteURL -like "*/sites/*") {

            $apiUri = "https://graph.microsoft.com/v1.0/sites/$($siteSplit[2]):/$($siteSplit[3])/$($siteSplit[4])/$($siteSplit[5])"
        
        }
        else {
            $apiUri = "https://graph.microsoft.com/v1.0/sites/$($siteSplit[2]):/$($siteSplit[3])"
        }
        $Site = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)

        $apiUri = "https://graph.microsoft.com/v1.0/sites/$($Site.id)/drives"
        [array]$Drives = RunQueryandEnumerateResults -token $token -apiUri $apiUri

        $Drive = $Drives | ? { $_.name -eq $LibraryName }

    }
    else {
        write-host "Processing same drive, continue..."
    }

    $apiUri = "https://graph.microsoft.com/v1.0/drives/$($drive.id)/root:$filepath"

    $item = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)

    $apiUri = "https://graph.microsoft.com/v1.0/drives/$($drive.id)/items/$($item.id)/permissions"

    $permissions = RunQueryandEnumerateResults -apiUri $apiUri -token $token

    foreach ($permission in $permissions) {

        $RequestBody = @"
        {
            "roles": [
                "read"
            ]
        }
"@


        $apiUri = "https://graph.microsoft.com/v1.0/drives/$($drive.id)/items/$($item.id)/permissions/$($permission.id)"
        #Run Graph Query
        Do {

            Try {
                (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)" } -Uri $apiUri -Method Patch -ContentType 'application/json' -Body $RequestBody)
                $HeaderCheck = $null
            }

            catch {
                $HeaderCheck = $null
                $HeaderCheck = $error[0]
                if ($HeaderCheck.Exception.Response.StatusCode -eq 429) {
                    write-host "Throttling encountered accessing $Apiuri, backing off temporarily" -ForegroundColor Yellow
                    start-sleep -Seconds 30
                    write-host "Resuming.." -ForegroundColor Yellow

                }else{
                    write-host "Could not update permission entry below on file $($Filepath): `n $($permission | convertto-json)" -ForegroundColor yellow
                }
            }

        }Until($HeaderCheck.Exception.Response.StatusCode -ne 429)

    }
}
