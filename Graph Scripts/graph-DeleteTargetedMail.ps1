<##Author: Sean McAvinue
##Details: Graph / PowerShell Script to target and delete specific emails from user mailboxes
##          Please fully read and test any scripts before running in your production environment!
        .SYNOPSIS
        Lists mails based on date filter and allows selection and deletion of one or multiple mails

        .PARAMETER Mailbox
        User Principal Name of target mailbox

        .PARAMETER StartDate
        Date to delete mails up to

        .PARAMETER EndDate
        Date to delete mails up to

        .PARAMETER ClientID
        Application (Client) ID of the App Registration

        .PARAMETER ClientSecret
        Client Secret from the App Registration

        .PARAMETER TenantID
        Directory (Tenant) ID of the Azure AD Tenant

        .EXAMPLE
        .\graph-DeleteTargetedMail.ps1 -Mailbox "adminseanmc@adminseanmc.com" -ClientSecret $clientSecret -ClientID $clientID -TenantID $tenantID -StartDate "2021-01-30T00:00:00Z" -EndDate "2021-06-01T00:00:00Z"
        
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
    $Mailbox,
    [parameter(Mandatory = $true)]
    [String]
    $StartDate,
    [parameter(Mandatory = $true)]
    [String]
    $EndDate

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

function DeleteMail {
    <#
.SYNOPSIS
Deltes mail from mailbox

.PARAMETER mail
mail ID

.PARAMETER token
Access token

.PARAMETER mailbox
Users UPN

#>
    Param(
        [parameter(Mandatory = $true)]
        $mail,
        [parameter(Mandatory = $true)]
        $token,
        [parameter(Mandatory = $true)]
        $mailbox

    )

    $Apiuri = "https://graph.microsoft.com/v1.0/users/$mailbox/messages/$mail/move"

    $Destination = @"
    {
        "destinationId": "recoverableitemspurges"
    }
"@

    (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -ContentType 'application/json' -Body $destination -Uri $apiUri -Method Post)

}


$token = GetGraphToken -ClientSecret $ClientSecret -ClientID $ClientID -TenantID $TenantID

$Apiuri = "https://graph.microsoft.com/v1.0/users/$mailbox/messages?`$filter=receivedDateTime le $EndDate and receivedDateTime gt $StartDate"

write-host "checking Mails via: $Apiuri"
Try {
    $results = RunQueryandEnumerateResults -apiUri $apiuri -token $token  | select-object  -property subject, receiveddatetime, @{name = "FromAddress"; Expression = { $_.from.emailaddress.address } }, id | Out-GridView -PassThru
}
catch {
    throw "Error searching for mails with given paramaters: $($_.Exception.Message)"
    break
}


write-host "Found $($results.count) mails"

foreach ($mail in $Results) {

    write-host "Processing $($mail.subject)"

    DeleteMail -token $token -mail $mail.id -mailbox $mailbox
}