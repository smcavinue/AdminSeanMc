<#
    .SYNOPSIS
    This Script creates a new guest user invitation and optionally sends the invitation or prints the redemption URL to the screen
    
    .PARAMETER UserDisplayName
    -UserDisplayName is the display name for the guest account

    .PARAMETER UserEmail
    -UserEmail is the email address of the requested user

    .PARAMETER UserMessage
    -UserMessage is the custom message to present to the user, can be used when sending intivation automatically
    
    .PARAMETER clientSecret
    -is the app registration client secret

    .PARAMETER clientID
    -is the app clientID

    .PARAMETER tenantID
    -is the directory ID of the tenancy
    
    .PARAMETER RedirectURL
    -A URL to redrect to after the invitation is redeemed

    .PARAMETER SendInvite
    -Define if the invitation should be sent ($true) or URL returned ($false)
    #>
Param(
    [parameter(Mandatory = $true)]
    [String]
    $UserDisplayName,
    [parameter(Mandatory = $true)]
    [String]
    $UserEmail,
    [parameter(Mandatory = $false)]
    [String]
    $UserMessage = " ",
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
    $RedirectURL = "https://myapps.microsoft.com",
    [parameter(Mandatory = $true)]
    [Boolean]
    $SendInvite

)

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

function CreateInvitation {
    <#
    .SYNOPSIS
    Short description
    
    .DESCRIPTION
    Long description
    
    .PARAMETER UserEmail
    Parameter description
    
    .PARAMETER RedirectURL
    Parameter description
    
    .PARAMETER SendInvite
    Parameter description
    
    .EXAMPLE
    An example
    
    .NOTES
    General notes
    #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $UserDisplayName,
        [parameter(Mandatory = $true)]
        [String]
        $UserEmail,
        [parameter(Mandatory = $true)]
        [String]
        $RedirectURL,
        [parameter(Mandatory = $true)]
        [Boolean]
        $SendInvite,
        [parameter(Mandatory = $true)]
        [String]
        $UserMessage

    )

    write-host "Creating User Invitation with the following settings:"
    $InvitationObject = @"
    {
        "invitedUserDisplayName": "$UserDisplayName",
        "invitedUserEmailAddress": "$userEmail",
        "sendInvitationMessage": "$SendInvite",
        "inviteRedirectUrl": "$RedirectURL",
        "invitedUserMessageInfo": {
            "customizedMessageBody": "$UserMessage"
        }
    }
"@
    $InvitationObject | out-host
    return $InvitationObject

}
    

$token = GetGraphToken -ClientSecret $ClientSecret -ClientID $ClientID -TenantID $TenantID

$apiUri = 'https://graph.microsoft.com/beta/invitations/'
$body = CreateInvitation -UserDisplayName $UserDisplayName -UserEmail $UserEmail -RedirectURL $RedirectURL -SendInvite $SendInvite -UserMessage $UserMessage
    
$invitation = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)" } -Uri $apiUri -Method Post -ContentType 'application/json' -Body $body)

if($sendinvite){
    write-host "Invitation has been sent to $useremail"
}else{

write-host "Invitation Redemption URL is: $($invitation.inviteRedeemUrl)"

}
