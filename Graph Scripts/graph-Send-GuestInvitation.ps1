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

Function SendGuestInvitation {
    <#
    .SYNOPSIS
    This function gets Graph Token from the GetGraphToken Function and uses it to request a new guest user
    
    .PARAMETER UserEmail
    -UserEmail is the email address of the requested user
    
    .PARAMETER clientSecret
    -is the app registration client secret

    .PARAMETER clientID
    -is the app clientID

    .PARAMETER tenantID
    -is the directory ID of the tenancy
    
    .PARAMETER tenantID
    -A URL to redrect to after the invitation is redeemed
    #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $UserEmail,
        [parameter(Mandatory = $true)]
        [String]
        $ClientSecret,
        [parameter(Mandatory = $true)]
        [String]
        $ClientID,
        [parameter(Mandatory = $true)]
        [String]
        $TenantID,
        [parameter(Mandatory = $false)]
        [String]
        $RedirectURL = "https://myapps.microsoft.com"

    )

    $token = GetGraphToken -ClientSecret $ClientSecret -ClientID $ClientID -TenantID $TenantID

    $apiUri = 'https://graph.microsoft.com/beta/invitations/'
    $body = "{'invitedUserEmailAddress': '$UserEmail','inviteRedirectUrl': '$RedirectURL'}"
    
    $invitation = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)" } -Uri $apiUri -Method Post -ContentType 'application/json' -Body $body)

    Return $invitation
}


