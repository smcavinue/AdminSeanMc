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

Install-Module MSAL.PS <br>
Install-Module ImportExcel <br>
Install-Module ExchangeOnlineManagement <br>
Install-Module AzureAD <br>

### Files and folders
1. C:\TEMP
2. Prepare-TenantAssessment.ps1
3. Perform-TenantAssessment.ps1
4. TenantAssessment-Template.xlsx

### Permissions
You need global admin permission in the tenant to create the App
You might need local admin permission to run the scripts (bypass, certificates, etc)

## Order of Execution

The script is not signed, so it might be required to bypass your execution policies

Set-ExecutionPolicy -ExecutionPolicy Unrestricted
### Copy files
1. Copy the files  <br>
-> Perform-TenantAssessment.ps1  <br>
-> Prepare-TenantAssessment.ps1  <br>
-> TenantAssessment-Template.xlsx <br>
to a local folder

2. If not exist, create the C:\Temp directory for the resulting files

### Prepare the Assessment
3. with local administrative permissions: run the script Prepare-TenantAssessment.ps1 (will create the App for the Assessment)
4. write down the ClientID, TenantID and CertificateThumbprint values !!! IMPORTANT FOR SECOND COMMAND !!!


### Perform the Assessment
5. run the script Perform-TenantAssessment.ps1 (does the magic)
   <br>
---> for example: .\Perform-TenantAssessment.ps1 -ClientID 38c24985-933a-46dc-90e6-5db54c777ef2 -TenantID 3548c4da-03e2-4e68-abbc-fcd68945d257 -CertificateThumbPrint F7D9B03E7DE24090D49B4CB974D74561CF375F9
7. open the Excel File created under C:\Temp
8. Have fun analyzing

## Open Features:
The assessment is not fully complete, of course (and will never be) , some more information can be added also using the already used technologies..
### inventory of public folders
- List of public folders
- Size of the PF
- Mailenabled
- reachable from external organizations
- permissions
- some more
- We need an additional excel sheet to do this
