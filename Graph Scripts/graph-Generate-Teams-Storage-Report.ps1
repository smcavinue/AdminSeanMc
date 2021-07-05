<##Author: Sean McAvinue
##Details: Graph / PowerShell Script to generate a report of Teams storage paths including Private Channels
##          Please fully read and test any scripts before running in your production environment!
        .SYNOPSIS
        Generates a CSV Report of Teams Storage including Private channels

        .PARAMETER ClientID
        Application (Client) ID of the App Registration

        .PARAMETER ClientSecret
        Client Secret from the App Registration

        .PARAMETER TenantID
        Directory (Tenant) ID of the Azure AD Tenant

        .PARAMETER CSVPath
        Path and name of the export CSV

        .EXAMPLE
        .\graph-Generate-Teams-Storage-Report.ps1  -ClientSecret $clientSecret -ClientID $clientID -TenantID $tenantID
        
        .Notes
        For similar scripts check out the links below
        
            Blog: https://seanmcavinue.net
            GitHub: https://github.com/smcavinue
            Twitter: @Sean_McAvinue
            Linkedin: https://www.linkedin.com/in/sean-mcavinue-4a058874/


    #>

##

Param(
    [parameter(Mandatory = $true)]
    [String]
    $ClientSecret,
    [parameter(Mandatory = $true)]
    [String]
    $ClientID,
    [parameter(Mandatory = $true)]
    [String]
    $TenantID,
    [parameter(Mandatory = $true)]
    [String]
    $CSVPath
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
    try {
        $tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
    }
    catch {

        throw "Error Retriving Access Token, verify provided details"
        break
    }
    # Access Token
    $token = ($tokenRequest.Content | ConvertFrom-Json).access_token
    
    #Returns token
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
    $Results = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)
    #Output Results for debug checking
    #write-host $results

    #Begin populating results
    $ResultsValue = $Results.value

    #If there is a next page, query the next page until there are no more pages and append results to existing set
    if ($results."@odata.nextLink" -ne $null) {
        #
        $NextPageUri = $results."@odata.nextLink"
        ##While there is a next page, query it and loop, append results
        While ($NextPageUri -ne $null) {
            write-host enumerating pages -ForegroundColor yellow
            $NextPageRequest = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $NextPageURI -Method Get)
            $NxtPageData = $NextPageRequest.Value
            $NextPageUri = $NextPageRequest."@odata.nextLink"
            $ResultsValue = $ResultsValue + $NxtPageData
        }
    }

    ##Return completed results
    return $ResultsValue

    
}


##Generate an access token
$token = GetGraphToken -clientId $ClientID -clientSecret $ClientSecret -tenantId $tenantID

##Build Request to get all groups with a Team provisioned
$apiuri = "https://graph.microsoft.com/beta/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"

##List all Teams
$Teams = RunQueryandEnumerateResults -apiUri $apiUri -token $token

foreach ($team in $teams) {

    ##Build Request to get Team site
    $apiUri = "https://graph.microsoft.com/v1.0/groups/$($Team.id)/drive"
    ##Get Storage Details of Team
    $Drive = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)

    ##Build Top Level Group Entry
    $ExportObject = [PSCustomObject]@{
        "Team ID"              = $Team.id
        "Team Name"            = $Team.DisplayName
        "Channel Name"         = "N/A"
        "Channel Type"         = "N/A"
        "SharePoint URL"       = $Drive.webUrl
        "Storage Used (Bytes)" = $Drive.quota.used
    }

    ##Export Team Size to Report
    $ExportObject | export-csv $csvPath -NoClobber -NoTypeInformation -Append

    ##Build Request to list Team Channels
    $apiuri = "https://graph.microsoft.com/v1.0/teams/$($Team.id)/channels"

    ##List Team Channels
    $Channels = RunQueryandEnumerateResults -apiUri $apiUri -token $token

    foreach ($Channel in $Channels) {

        ##Build Private channel files folder query
        $apiUri = "https://graph.microsoft.com/beta/teams/$($Team.id)/channels/$($Channel.id)/filesfolder"

        try {
            ##Get Storage Details of Channel
            $Drive = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)
        }
        catch {

            write-host "Channel files not provisioned for $apiuri" 
            #$team.displayName | out-file c:\temp\channelsnotfound.csv -append

        }
        ##Build Private Channel Entry
        $ExportObject = [PSCustomObject]@{
            "Team ID"              = $Team.id
            "Team Name"            = $Team.DisplayName
            "Channel Name"         = $Channel.displayName
            "Channel Type"         = $Channel.membershipType
            "SharePoint URL"       = $Drive.webUrl
            "Storage Used (Bytes)" = $Drive.size
        }
    
        ##Export Team Size to Report
        $ExportObject | export-csv $CSVPath -NoClobber -NoTypeInformation -Append

    }
}