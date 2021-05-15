$tempfile = "c:\temp\microsoft.identitymodel.clients.activedirectory.3.19.8.zip"
$tempfolder = "c:\temp\microsoft.identitymodel.clients.activedirectory.3.19.8"
$URI = "https://www.nuget.org/api/v2/package/Microsoft.IdentityModel.Clients.ActiveDirectory/3.19.8"
$ModuleFile = "C:\temp\microsoft.identitymodel.clients.activedirectory.3.19.8\lib\net45\SetPlannerTenantSettings.psm1"
$ManifestFile = "C:\temp\microsoft.identitymodel.clients.activedirectory.3.19.8\lib\net45\SetPlannerTenantSettings.psd1"



if(!(Test-Path "C:\temp")){

    new-item -ItemType directory -Path c:\temp -Force 

}
invoke-webrequest -uri $uri -outfile $tempfile

Expand-Archive -LiteralPath $tempfile -DestinationPath $tempfolder



$modulecontents = {
    function Connect-AAD (){
<#
.Synopsis
(Private to module) Attempts to obtain a token from AAD.
.Description
This function attempts to obtain a token from Azure Active Directory.
.example
$authorizationContext = Connect-AAD
#>

    $authUrl = "https://login.microsoftonline.com/common" # Prod environment
    $resource = "https://tasks.office.com" # Prod environment
    $clientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c"

    $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
    $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Always"

    $authentiationContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authUrl, $False

    $authenticationResult = $authentiationContext.AcquireTokenAsync($resource, $clientId, $redirectUri, $platformParameters).Result
    return $authenticationResult
}

function Set-PlannerConfiguration
{
<#
.Synopsis
Configures tenant level settings for Microsoft Planner.
.Description
This cmdlet allows tenant administrators to change policies regarding availability of certain features in Microsoft Planner. The changes to settings take effect immediately. This cmdlet specifies the administrator preference on whether the feature should be available. The features can still be disabled due to Microsoft Planner behavior, at the discretion of Microsoft.
.Parameter Uri
The URL of the Tenant-Level Settings API for the Planner instance to control.
.Parameter AccessToken
A valid access token of a user with tenant-level administrator privileges.
.Parameter AllowCalendarSharing
If set to $false, disables creating iCalendar links from Microsoft Planner, and disables previously created iCalendar links.  If set to $true, enables creating iCalendar links from Microsoft Planner and re-enables any previously created iCalendar links.
.Parameter AllowTenantMoveWithDataLoss
If set to $true, allows the tenant to be moved to another Planner environment or region. This move will result in the tenant's existing Planner data being lost.
.Parameter AllowRosterCreation
If set to $true, allows the users of the tenant to create rosters as the container for a plan to facilitate ad-hoc collaboration. This setting does not restrict the use of existing roster contained plans.
.Parameter AllowPlannerMobilePushNotifications
If set to $true, allows the use of direct push mobile notifications in Tenant
.example

Set-PlannerConfiguration -AllowCalendarSharing $true

.example

Set-PlannerConfiguration -AllowTenantMoveWithDataLoss $true

.example

Set-PlannerConfiguration -AllowRosterCreation $false

.example

Set-PlannerConfiguration -AllowPlannerMobilePushNotifications $false
#>
    param(
        [ValidateNotNull()]
        [System.String]$Uri="https://tasks.office.com/taskAPI/tenantAdminSettings/Settings",

        [ValidateNotNullOrEmpty()]
        [Parameter(Mandatory=$false)][System.String]$AccessToken,
        [Parameter(Mandatory=$false, ValueFromPipeline=$true)][System.Boolean]$AllowCalendarSharing,
        [Parameter(Mandatory=$false, ValueFromPipeline=$true)][System.Boolean]$AllowTenantMoveWithDataLoss,
        [Parameter(Mandatory=$false, ValueFromPipeline=$true)][System.Boolean]$AllowRosterCreation,
        [Parameter(Mandatory=$false, ValueFromPipeline=$true)][System.Boolean]$AllowPlannerMobilePushNotifications
        )

    if(!($PSBoundParameters.ContainsKey("AccessToken"))){
        $authorizationContext = Connect-AAD
        $AccessToken = $authorizationContext.AccessTokenType.ToString() + ' ' +$authorizationContext.AccessToken
    }

    $flags = @{}

    if($PSBoundParameters.ContainsKey("AllowCalendarSharing")){
        $flags.Add("allowCalendarSharing", $AllowCalendarSharing);
    }

    if($PSBoundParameters.ContainsKey("AllowTenantMoveWithDataLoss")){
        $flags.Add("allowTenantMoveWithDataLoss", $AllowTenantMoveWithDataLoss);
    }

    if($PSBoundParameters.ContainsKey("AllowRosterCreation")){
        $flags.Add("allowRosterCreation", $AllowRosterCreation);
    }

    if($PSBoundParameters.ContainsKey("AllowPlannerMobilePushNotifications")){
        $flags.Add("allowPlannerMobilePushNotifications", $AllowPlannerMobilePushNotifications)
    }

    $propertyCount = $flags | Select-Object -ExpandProperty Count

    if($propertyCount -eq 0) {
        Throw "No properties were set."
    }

    $requestBody = $flags | ConvertTo-Json

    Invoke-RestMethod -ContentType "application/json;odata.metadata=full" -Headers @{"Accept"="application/json"; "Authorization"=$AccessToken; "Accept-Charset"="UTF-8"; "OData-Version"="4.0;NetFx"; "OData-MaxVersion"="4.0;NetFx"} -Method PATCH -Body $requestBody $Uri
}

function Get-PlannerConfiguration
{
<#
.Synopsis
Retrieves tenant level settings for Microsoft Planner.
.Description
This cmdlet allows users and tenant administrators to retrieve policy preferences set by the tenant administrator regarding availability of certain features in Microsoft Planner.  While a feature may be permitted by a tenant administrator's preference, features can still be disabled due to Microsoft Planner behavior, at the discretion of Microsoft.
.Parameter Uri
The URL of the Tenant-Level Settings API for the Planner instance to retrieve.
.Parameter AccessToken
A valid access token of a user with tenant-level administrator privileges.

.example

Get-PlannerConfiguration
#>
    param(
        [ValidateNotNull()]
        [System.String]$Uri="https://tasks.office.com/taskAPI/tenantAdminSettings/Settings",

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [System.String]$AccessToken
        )

    if(!($PSBoundParameters.ContainsKey("AccessToken"))){
        $authorizationContext = Connect-AAD
        $accessToken = $authorizationContext.AccessTokenType.ToString() + ' ' +$authorizationContext.AccessToken
    }

    $response = Invoke-RestMethod -ContentType "application/json;odata.metadata=full" -Headers @{"Accept"="application/json"; "Authorization"=$AccessToken; "Accept-Charset"="UTF-8"; "OData-Version"="4.0;NetFx"; "OData-MaxVersion"="4.0;NetFx"} -Method GET $Uri
    $result = New-Object PSObject -Property @{
        "AllowCalendarSharing" = $response.allowCalendarSharing
        "AllowTenantMoveWithDataLoss" = $response.allowTenantMoveWithDataLoss
        "AllowRosterCreation" = $response.allowRosterCreation
        "AllowPlannerMobilePushNotifications" = $response.allowPlannerMobilePushNotifications
    }

    return $result
}

Export-ModuleMember -Function Get-PlannerConfiguration, Set-PlannerConfiguration
}

$ManifestContents = {
#
# Module manifest for module 'SetTenantSettings'
#
# Generated by: Microsoft Corporation
#
# Generated on: 12/17/2017
#

@{


    RootModule = 'SetPlannerTenantSettings.psm1' 

    ModuleVersion = '1.0' 

    CompatiblePSEditions = @() 

    GUID = '6250c644-4898-480c-8e0b-bd3ebdf246ca' 

    Author = 'Microsoft Corporation' 

    CompanyName = 'Microsoft Corporation' 

    Copyright = '(c) 2017 Microsoft Corporation. All rights reserved.' 

    Description = 'Planner Tenant Settings client' 

    PowerShellVersion = '' 
 
    PowerShellHostName = '' 

    PowerShellHostVersion = '' 

    DotNetFrameworkVersion = '' 

    CLRVersion = '' 

    ProcessorArchitecture = '' 

    RequiredModules = @() 

    RequiredAssemblies = @("Microsoft.IdentityModel.Clients.ActiveDirectory.dll","Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll") 

    ScriptsToProcess = @() 

    TypesToProcess = @() 

    FormatsToProcess = @() 

    NestedModules = @() 

    FunctionsToExport = @("Get-PlannerConfiguration", "Set-PlannerConfiguration") 

    CmdletsToExport = @() 

    VariablesToExport = '*' 

    AliasesToExport = @() 

    DscResourcesToExport = @() 

    ModuleList = @() 

    FileList = @("SetTenantSettings.psm1") 

    PrivateData = @{ PSData = @{ 

    Tags = @() 

    LicenseUri = '' 

    ProjectUri = '' 

    IconUri = '' 

    ReleaseNotes = '' } 

    } 

    HelpInfoURI = '' 

    DefaultCommandPrefix = '' }

}

$modulecontents | out-file $ModuleFile -Width 4096
$ManifestContents | out-file $ManifestFile -Width 4096

