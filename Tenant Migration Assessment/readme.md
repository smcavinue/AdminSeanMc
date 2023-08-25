# Office 365 migration plan Powershell assessment prerequisites
## Original source
from https://practical365.com/office-365-migration-plan-assessment/

## Requirements
### Modules
The Office 365 migration plan report requires four PowerShell Modules

1. MSAL.PS authenticates with the Microsoft Graph API and obtains an Access Token
2. ImportExcel exports the report to Excel
3. ExchangeOnlineManagement connects to Exchange Online to get information that is not exposed via the Graph API
4. AzureAD (or AzureADPreview) sets up the required app registration and permissions

If you don't already have the required modules installed on a workstation, you can install them from the PS Gallery by running these commands:

Install-Module MSAL.PS
Install-Module ImportExcel
Install-Module ExchangeOnlineManagement
Install-Module AzureAD

### Files and folders
1. C:\TEMP
2. Prepare-TenantAssessment.ps1
3. Perform-TenantAssessment.ps1
4. TenantAssessment-Template.xlsx

### Permissions
You need global admin permission in the tenant to create the App
You need local admin permission to run the scripts (bypass, certificates, etc)

## Order of Execution

The script is not signed, so it might be required to bypass your execution policies

Set-ExecutionPolicy -ExecutionPolicy Unrestricted

1. Copy the files Perform-TenantAssessment.ps1, Prepare-TenantAssessment.ps1 and TenantAssessment-Template.xlsx to a local folder
2. If not exist, create the C:\Temp directory
3. with local administrative permissions: run the script Prepare-TenantAssessment.ps1 (will create the App for the Assessment)
4. write down the ClientID, TenantID and CertificateThumbprint values !!! IMPORTANT FOR SECOND COMMAND !!!
5. run the script Perform-TenantAssessment.ps1 (does the magic)
---> for example: .\Perform-TenantAssessment.ps1 -ClientID 38c24985-933a-46dc-90e6-5db54c777ef2 -TenantID 3548c4da-03e2-4e68-abbc-fcd68945d257 -CertificateThumbPrint F7D9B03E7DE24090D49B4CB974D74561CF375F9
7. open the Excel File created under C:\Temp
8. Have fun analyzing
