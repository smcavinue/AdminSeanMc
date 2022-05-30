<##Author: Sean McAvinue
##Details: PowerShell Script to Configure an Application Registration with the appropriate permissions to run Perform-TenantAssessment.ps1
##          Please fully read and test any scripts before running in your production environment!
        .SYNOPSIS
        Creates an app reg with the appropriate permissions to run the Conditional Access assessment script and uploads a self signed certificate

        .DESCRIPTION
        Connects to Azure AD and provisions an app reg with the appropriate permissions

        .Notes
        For similar scripts check out the links below
        
            Blog: https://seanmcavinue.net
            GitHub: https://github.com/smcavinue
            Twitter: @Sean_McAvinue
            Linkedin: https://www.linkedin.com/in/sean-mcavinue-4a058874/


    #>

##
#Requires -modules azuread
Param(
    [Parameter(Mandatory = $false)]
    [Switch]$UseClientSecret
)
function New-AadApplicationCertificate {
    [CmdletBinding(DefaultParameterSetName = 'DefaultSet')]
    Param(
        [Parameter(mandatory = $true)]
        [string]$CertificatePassword,

        [Parameter(mandatory = $true, ParameterSetName = 'ClientIdSet')]
        [string]$ClientId,

        [string]$CertificateName,

        [Parameter(mandatory = $false, ParameterSetName = 'ClientIdSet')]
        [switch]$AddToApplication
    )
    ##Function source: https://www.powershellgallery.com/packages/AadSupportPreview/0.3.8/Content/functions%5CNew-AadApplicationCertificate.ps1

    # Create self-signed Cert
    $notAfter = (Get-Date).AddYears(2)

    try {
        $cert = (New-SelfSignedCertificate -DnsName "ConditionalAccessAssessment" -CertStoreLocation "cert:\currentuser\My" -KeyExportPolicy Exportable -Provider "Microsoft Enhanced RSA and AES Cryptographic Provider" -NotAfter $notAfter)
        
        #Write-Verbose "Cert Hash: $($cert.GetCertHash())"
        #Write-Verbose "Cert Thumbprint: $($cert.Thumbprint)"
    }

    catch {
        Write-Error "ERROR. Probably need to run as Administrator."
        Write-host $_
        return
    }

    if ($AddToApplication) {
        $AppObjectId = $app.ObjectId
        $KeyValue = [System.Convert]::ToBase64String($cert.GetRawCertData())
        New-AzureADApplicationKeyCredential -ObjectId $appReg.ObjectId -Type AsymmetricX509Cert -Usage Verify -Value $KeyValue | out-null

    }
    Return $cert.Thumbprint
}

##Declare Variables
##Monitors connection attempt
$connected = $false
##Name of the app
$appName = "Conditional Access Assessment"
##The URI of the app - set to localhost
$appURI = @("https://localhost")
##Contain settings of the app reg
$appReg = $null
##Consent URL
$ConsentURl = "https://login.microsoftonline.com/{tenant-id}/adminconsent?client_id={client-id}"
##Tenant ID
$TenantID = $null

##Attempt Azure AD connection until successful
while ($connected -eq $false) {
    Try {
        Connect-AzureAD -ErrorAction stop
        $connected = $true
    }
    catch {
        Write-Host "Error connecting to Azure AD: `n$($error[0])`n Try again..." -ForegroundColor Red
        $connected = $false
    }
}

##Create Resource Access Variable
Try {
    $Permissions = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
    ##Declare Application Permission - Reference here: https://docs.microsoft.com/en-us/graph/permissions-reference
    $permList = @(
        "5b567255-7703-4780-807c-7be8301ae99b",
        "246dd0d5-5bd0-4def-940b-0421030a5b68",
        "df021288-bdef-4463-88db-98f22de89214",
        "7ab1d382-f21e-4acd-a863-ba3e13f7da61",
        "37730810-e9ba-4e46-b07e-8ca78d182097"
    )

    $permArray = @()
    foreach ($perm in $permList) {
        ##Create perm
        $permArray += New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList $perm, "Role"
        
    }
    ##Add permission list to object
    $permissions.ResourceAccess = $permArray
    $permissions.ResourceAppId = "00000003-0000-0000-c000-000000000000"
}
Catch {

    Write-Host "Error preparing script: `n$($error[0])`nCheck Prerequisites`nExiting..." -ForegroundColor Red
    pause
    exit

}


##Check for existing app reg with the same name
$AppReg = Get-AzureADApplication -Filter "DisplayName eq '$($appName)'"  -ErrorAction SilentlyContinue

##If the app reg already exists, do nothing
if ($appReg) {
    write-host "App already exists - Please delete the existing 'Conditional Access Assessment' app from Azure AD and rerun the preparation script to recreate, exiting" -ForegroundColor yellow
    Pause
    exit
}
else {

    Try {
        ##Create the new App Reg
        $appReg = New-AzureADApplication -DisplayName $appName -ReplyUrls $appURI -ErrorAction Stop -RequiredResourceAccess $Permissions
        
        Write-Host "Waiting for app to provision..."
        start-sleep -Seconds 20
    }
    catch {
        Write-Host "Error creating new app reg: `n$($error[0])`n Exiting..." -ForegroundColor Red
        pause
        exit
    }

}

If ($UseClientSecret) {
    $appSecret = New-AzureADApplicationPasswordCredential -ObjectId $appReg.objectId -CustomKeyIdentifier ((get-date).ToString().Replace('/', '')) -StartDate (get-date) -EndDate ((get-date).AddDays(1))
}
else {
    $Thumbprint = New-AadApplicationCertificate -ClientId $appReg.AppId -CertificatePassword "T3mPP@£6hnhskke!!!" -AddToApplication -certificatename "Tenant Assessment Certificate"
}
##Get tenant ID
$tenantID = (Get-AzureADTenantDetail).objectid
##Update Consent URL
$ConsentURl = $ConsentURl.replace('{tenant-id}', $TenantID)
$ConsentURl = $ConsentURl.replace('{client-id}', $appReg.AppId)

write-host "Consent page will appear, don't forget to log in as admin to grant consent!" -ForegroundColor Yellow
Start-Process $ConsentURl
if ( $UseClientSecret) {
    Write-Host "The below details can be used to run the assessment, take note of them and press any button to clear the window.`nTenant ID: $tenantID`nClient ID: $($appReg.appID)`nClient Secret: $($appSecret.value)" -ForegroundColor Green
}else {
    Write-Host "The below details can be used to run the assessment, take note of them and press any button to clear the window.`nTenant ID: $tenantID`nClient ID: $($appReg.appID)`nCertificate Thumbprint: $thumbprint" -ForegroundColor Green   
}
Pause
clear
