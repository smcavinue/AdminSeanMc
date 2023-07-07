# Office 365 migration plan Powershell assessment prerequisites
from https://practical365.com/office-365-migration-plan-assessment/
## Requirements
### Modules
The Office 365 migration plan report requires four PowerShell Modules

1. MASL.PS authenticates with the Microsoft Graph API and obtains an Access Token
2. ImportExcel exports the report to Excel
3. ExchangeOnlineManagement connects to Exchange Online to get information that is not exposed via the Graph API
4. AzureAD (or AzureADPreview) sets up the required app registration and permissions

If you don't already have the required modules installed on a workstation, you can install them from the PS Gallery by running these commands:

Install-Module MSAL.PS
Install-Module ImportExcel
Install-Module ExchangeOnlineManagement
Install-Module AzureAD
### Permissions
You need global admin permission in the tenant to create the App
You need local admin permission to rn the scripts (bypass, certificates, etc)

## Order of Execution

The script is not signed, so it might be required to bypass your execution policies

Set-ExecutionPolicy -ExecutionPolicy Unrestricted

1. Copy the files Perform-TenantAssessment.ps1, Prepare-TenantAssessment.ps1 and TenantAssessment-Template.xlsx to a local folder
2. If not exist, create the C:\Temp directory
3. run the script Prepare-TenantAssessment.ps1 (will create the App for the Assessment)
4. write down the ClientID, TenantID and CertificateThumbprint values
5. run the script Perform-TenantAssessment.ps1 (does the magic)
6. open the Excel File created under C:\Temp
