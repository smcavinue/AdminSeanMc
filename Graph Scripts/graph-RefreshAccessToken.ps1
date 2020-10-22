function RefreshAccessToken{
    <#
    .SYNOPSIS
    Refreshes an access token based on refresh token

    .RETURNS
    Returns a refreshed access token
    
    .PARAMETER Token
    -Token is the existing token

    .PARAMETER tenantID
    -This is the tenant ID eg. domain.onmicrosoft.com

    .PARAMETER ClientID
    -This is the app reg client ID

    .PARAMETER Secret
    -This is the client secret

    .PARAMETER Scope
    -A comma delimited list of access scope, default is: "Group.ReadWrite.All,User.ReadWrite.All"
    
    #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $Token,
        [parameter(Mandatory = $true)]
        [String]
        $tenantID,
        [parameter(Mandatory = $true)]
        [String]
        $ClientID,
        [parameter(Mandatory = $false)]
        [String]
        $Scope = "Group.ReadWrite.All,User.ReadWrite.All",
        [parameter(Mandatory = $true)]
        [String]
        $Secret
    )

$ScopeFixup = $Scope.replace(',','%20')
$apiUri = "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token"
$body = "client_id=$ClientID&scope=$ScopeFixup&refresh_token=$Token&redirect_uri=http%3A%2F%2Flocalhost%2F&grant_type=refresh_token&client_secret=$Secret"
write-verbose $body -Verbose
$Refreshedtoken = (Invoke-RestMethod -Uri $apiUri -Method Post -ContentType 'application/x-www-form-urlencoded' -body $body  )

return $Refreshedtoken

}
