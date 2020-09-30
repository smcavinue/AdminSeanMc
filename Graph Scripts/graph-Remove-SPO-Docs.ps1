#Author: Sean McAvinue
#Description: Removes files that match the search criteria. This script is to be used as an example, not tested for production
function GetGraphToken {
    # Azure AD OAuth Application Token for Graph API
    # Get OAuth token for a AAD Application (returned as $token)
     
    # Application (client) ID, tenant ID and secret
    $clientId = "<ClientID>"
    $tenantId = "<TenantID>"
    $clientSecret = "<ClientSecret>"
    
    
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
  
##Tenant Name##
$TenantID = "TenantIdentifier"

##Search Term#
$SearchTerm = "String in file name"

$token = getgraphtoken
###Get Sites##
$sites = import-csv sitestoCheck.csv

##Loop Through Sites
foreach ($site in $sites) {
    
    ##Trim Site URL
    $SiteTrimmed = $site.url.split('/')
    $Entry = ($sitetrimmed[$sitetrimmed.count - 1])
    
    ##Query Site to get Site ID
    $apiUrl = 'https://graph.microsoft.com/beta/sites/' + $tenantId + '.sharepoint.com:/sites/' + $entry + ':/drives'
    $Data = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUrl -Method Get)
    
    ##Get the correct site ID 
    $SiteID = ($Data.value | ? { $_.name -like "Documents" }).ID
    
    ##Query Files in the site
    $apiUrl = ("https://graph.microsoft.com/v1.0/drives/" + $siteid + "/root/children")
    $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $token" } -Uri $apiUrl -Method Get
    
    ##Filter for only files matching syntax
    $FileToRemove = ($Data.value | ? { $_.name -like "*" + $SearchTerm + "*" })
    
    ##Check if Array, in case there are multiple files
    IF ($FileToRemove -is [array]) {
        ##Loop through Files if Array
        foreach ($File in $FileToRemove) {
    
            ##Write to screen and then delete
            write-host "Found" $File.name -ForegroundColor yellow 
            $apiUrl = ("https://graph.microsoft.com/v1.0/drives/" + $siteid + "/items/" + $File.id)
            $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $token" } -Uri $apiUrl -Method Delete
        }
    }##Else process single file
    elseif ($FileToRemove) {
    
        ##Write to screen and then delete
        write-host "Found" $FileToRemove.name -ForegroundColor green   
        $apiUrl = ("https://graph.microsoft.com/v1.0/drives/" + $siteid + "/items/" + $Filetoremove.id)
        $Data = Invoke-RestMethod -Headers @{Authorization = "Bearer $token" } -Uri $apiUrl -Method Delete
    
    }
    
    ##Tidy up Files and  FilesToRemove Variable
    remove-variable FiletoRemove, files
}
    
    
    
    