<##Author: Sean McAvinue
##Details: Graph / PowerShell Script t populate user contacts based on CSV input, 
##          Please fully read and test any scripts before running in your production environment!
        .SYNOPSIS
        Populates mail contacts into user mailboxes from a CSV

        .DESCRIPTION
        Creates a new mail contact for each entry in the input CSV in the target mailbox.

        .PARAMETER Mailbox
        User Principal Name of target mailbox

        .PARAMETER CSVPath
        Full path to the input CSV

        .PARAMETER ClientID
        Application (Client) ID of the App Registration

        .PARAMETER ClientSecret
        Client Secret from the App Registration

        .PARAMETER TenantID
        Directory (Tenant) ID of the Azure AD Tenant

        .EXAMPLE
        .\graph-PopulateContactsFromCSV.ps1 -Mailbox  $mailbox -ClientSecret $clientSecret -ClientID $clientID -TenantID $tenantID -CSVPath $csv
        
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
    $CSVPath

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

function ImportContact {
    <#
.SYNOPSIS
Imports contact into specified user mailbox

.DESCRIPTION
This function accepts an AAD token, user account and contact object and imports the contact into the users mailbox

.PARAMETER Mailbox
User Principal Name of target mailbox

.PARAMETER Contact
Contact object for processing

.PARAMETER Token
Access Token

        
#>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $Token,
        [parameter(Mandatory = $true)]
        [String]
        $Mailbox,
        [parameter(Mandatory = $true)]
        [PSCustomObject]
        $contact
    
    )
    write-host "contactcompanyname $($contact.companyname)"
    write-host $contact
    $ContactObject = @"
    {
        "assistantName": "$($contact.assistantName)",
        "businessHomePage": "$($contact.businessHomePage)",
        "businessPhones": [
            "$($contact.businessPhones)"
          ],
        "companyName": "$($contact.companyName)",
        "department": "$($contact.department)",
        "displayName": "$($contact.displayName)",
        "emailAddresses": [
            {
                "address": "$($contact.emailaddress)",
                "name": "$($contact.displayname)"
            }
        ],
        "givenName": "$($contact.givenname)",
        "jobTitle": "$($contact.jobTitle)",
        "middleName": "$($contact.middleName)",
        "nickName": "$($contact.nickName)",
        "profession": "$($contact.profession)",
        "personalNotes": "$($contact.personalNotes)",
        "surname": "$($contact.surname)",
        "title": "$($contact.title)",
        "fileAs": "$($contact.surname), $($contact.givenName)"
    }
"@


    write-host "contact object: $contactobject"

    $apiUri = "https://graph.microsoft.com/v1.0/users/$mailbox/contacts"
    write-host $apiuri
    Try {
        $NewContact = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -ContentType 'application/json' -Body $contactobject -Uri $apiUri -Method Post)
        return $NewContact
    }
    catch {
        throw "Error creating contact $($contact.emailaddress) for $mailbox $($_.Exception.Message)"
        break
    }
}

##MAIN##

##Try import CSV file
try {
    $Contacts = import-csv $CSVPath -ErrorAction stop
}
catch {
    throw "Error importing CSV: $($_.Exception.Message)"
    break
}

##Get Graph Token
Try {
    $Token = GetGraphToken -ClientSecret $ClientSecret -ClientID $ClientID -TenantID $TenantID
}
catch {
    throw "Error obtaining Token"
    break
}

##ProcessImports
foreach ($contact in $contacts) {

    $NewContact = ImportContact -Mailbox $mailbox -token $token -contact $contact

}

    
