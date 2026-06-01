#Requires -Modules "Az.Accounts", "Az.Resources", "Microsoft.Graph.Applications"
[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]$AutomationAccountName,
    [Parameter(Mandatory=$true)]
    [string]$Tenant,
    [Parameter(Mandatory=$true)]
    [string]$Subscription
)

$GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"

Connect-AzAccount -TenantId $Tenant -Subscription $Subscription  | Out-Null
Connect-MgGraph -TenantId $Tenant -Scopes "AppRoleAssignment.ReadWrite.All", "Application.Read.All" -NoWelcome

Write-Host "AZ context"
Get-AzContext | Format-List
Write-Host "MG context"
Get-MgContext | Format-List

$GraphPermissions = "Directory.Read.All", "Sites.Read.All", "Files.ReadWrite.All"
$AutomationMSI = (Get-AzADServicePrincipal -Filter "displayName eq '$AutomationAccountName'")
Write-Host "Assigning permissions to $AutomationAccountName ($($AutomationMSI.Id))"

$GraphServicePrincipal = Get-AzADServicePrincipal -Filter "appId eq '$GRAPH_APP_ID'"
$GraphAppRoles = $GraphServicePrincipal.AppRole | Where-Object {$_.Value -in $GraphPermissions -and $_.AllowedMemberType -contains "Application"}

if($GraphAppRoles.Count -ne $GraphPermissions.Count)
{
    Write-Warning "App roles found: $($GraphAppRoles)"
    throw "Some App Roles are not found on Graph API service principal"
}

foreach ($AppRole in $GraphAppRoles) {
    Write-Host "Assigning $($AppRole.Value) to $($AutomationMSI.DisplayName)"
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $AutomationMSI.Id -PrincipalId $AutomationMSI.Id -ResourceId $GraphServicePrincipal.Id -AppRoleId $AppRole.Id | Out-Null
}