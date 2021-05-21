<##Author: Sean McAvinue
##Details: Graph / PowerShell Script to list all Teams, Channels and tabs and export to CSV
##          Please fully read and test any scripts before running in your production environment!
        .SYNOPSIS
        Lists mails based on date filter and allows selection and deletion of one or multiple mails

        .PARAMETER ClientID
        Application (Client) ID of the App Registration

        .PARAMETER ClientSecret
        Client Secret from the App Registration

        .PARAMETER TenantID
        Directory (Tenant) ID of the Azure AD Tenant

        .PARAMETER CSVPathName
        The full path and name for CSV output

        .EXAMPLE
        .\graph-ListTeamsChannelsTabs.ps1  -ClientSecret $clientSecret -ClientID $clientID -TenantID $tenantID -CSVPathName c:\temp\Teamstabs.csv
        
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
    $CSVPathName

)

##FUNCTIONS##
function GetGraphToken {
    # Azure AD OAuth Application Token for Graph API
    # Get OAuth token for a AAD Application (returned as $token)
    <#
        .SYNOPSIS
        This function gets and returns a Graph Token using the provided details
    
        .PARAMETER clientSecret
        -is the app registration client secret
    
        .PARAMETER clientID
        -is the app clientID
    
        .PARAMETER tenantID
        -is the directory ID of the tenancy
        
        #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $ClientSecret,
        [parameter(Mandatory = $true)]
        [String]
        $ClientID,
        [parameter(Mandatory = $true)]
        [String]
        $TenantID
    
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

$token = GetGraphToken -ClientSecret $ClientSecret -ClientID $ClientID -TenantID $TenantID

$Apiuri = "https://graph.microsoft.com/v1.0/groups"
write-host $Apiuri
$Groups = RunQueryandEnumerateResults -token $token -apiUri $Apiuri

$Teams = ($groups | ?{$_.resourceProvisioningOptions -eq "Team"})

foreach($team in $teams){

    $apiuri = "https://graph.microsoft.com/v1.0/teams/$($team.id)/Channels"
    write-host $Apiuri
    $Channels = RunQueryandEnumerateResults -token $token -apiUri $apiuri
    
    foreach($channel in $channels){

        $apiuri = "https://graph.microsoft.com/v1.0/teams/$($team.id)/Channels/$($Channel.id)/tabs?`$expand=teamsApp"
        $tabs = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)

        foreach($tab in $tabs.value){

            $ExportObject = [pscustomobject]@{
                TeamId                = $team.id
                TeamDisplayName       = $team.displayName
                TeamIsArchived        = $team.isArchived
                TeamVisibility        = $team.visibility
                ChannelId             = $channel.id
                ChannelDisplayName    = $channel.DisplayName
                ChannelMemberShipType = $channel.membershipType
                TabId                 = $tab.id
                TabNameDisplayName    = $tab.DisplayName
                TeamsApp         = $tab.teamsApp.displayname
              }

            $exportobject | export-csv $CSVPathName -NoClobber -NoTypeInformation -Append


        }
    }

}