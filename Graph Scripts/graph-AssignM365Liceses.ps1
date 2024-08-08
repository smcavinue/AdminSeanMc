##Assign Microsoft 365 Licenses to Groups or Users Through PowerShell
# This script assigns Microsoft 365 licenses to groups or users through PowerShell.
# Author: Sean McAvinue
# Resources: https://seanmcavinue.net
# Example: Assign-M365Licenses.ps1

#Requires -modules microsoft.graph

Try {
    Connect-MgGraph -scopes "Directory.Read.All Organization.Read.All"
}
catch {
    Write-Host "Failed to connect to Microsoft Graph. Please check your credentials and try again."
    Write-Host $_.Exception.Message -ForegroundColor Red
    Exit
}

##Get All Licenses
[array]$LicensesAvailable = Get-MgSubscribedSku -All

$AvailableLicenseArray = @()
foreach ($LicenseAvailable in $LicensesAvailable) {

    $AvailableLicenses = [PSCustomObject]@{

        LicenseName         = $LicenseAvailable.SkuPartNumber
        LicenseId           = $LicenseAvailable.SkuId
        LicenseStatus       = $LicenseAvailable.capabilityStatus
        LicenseTotal        = $LicenseAvailable.PrepaidUnits.Enabled
        LicenseWithWarning  = $LicenseAvailable.PrepaidUnits.Warning
        LicensesDisabled    = $LicenseAvailable.PrepaidUnits.Suspended
        LicenseAssigned     = $LicenseAvailable.ConsumedUnits
        LicenseRemaining    = $LicenseAvailable.PrepaidUnits.Enabled - $LicenseAvailable.ConsumedUnits
        LicenseServicePlans = $LicenseAvailable.ServicePlans
    }
    $AvailableLicenseArray += $AvailableLicenses
}

[array]$LicenseSelection = $AvailableLicenseArray  | Out-GridView -PassThru -Title "Select Licenses to Assign"

If ($LicenseSelection) {

    ##Ask user if they want to disable Service Plans
    do {
        $DisableServicePlans = Read-Host "Do you want to disable any Service Plans? (Y/N)"
        if ($DisableServicePlans -ne "Y" -and $DisableServicePlans -ne "N") {
            Write-Host "Invalid Selection. Please select Y for Yes or N for No" -ForegroundColor Yellow
        }
    }until($DisableServicePlans -eq "Y" -or $DisableServicePlans -eq "N")

    If ($DisableServicePlans -eq "Y") {
    

        ##If Yes, Prompt user to select Service Plans to disable
        foreach ($SelectedLicense in $LicenseSelection) {

            [array]$ServicePlanSelection = $SelectedLicense.LicenseServicePlans | Out-GridView -PassThru -Title "Select Service Plans to Disable for $($SelectedLicense.LicenseName)"
            if ($ServicePlanSelection) {
                $SelectedLicense | Add-Member -MemberType NoteProperty -Name "DisabledServicePlans" -Value ($ServicePlanSelection.ServicePlanID) -Force 
            }
        }
    }
    ##Prompt user to select Groups or Users
    do {
        $UserOrGroup = Read-Host "Do you want to assign licenses to Users or Groups? (U/G)"
        if ($UserOrGroup -ne "U" -and $UserOrGroup -ne "G") {
            Write-Host "Invalid Selection. Please select U for Users or G for Groups" -ForegroundColor Yellow
        }
    }until($UserOrGroup -eq "U" -or $UserOrGroup -eq "G")

    ##If User Selected
    If ($UserOrGroup -eq "U") {
        [array]$Users = Get-MgUser -All | Out-GridView -PassThru -Title "Select Users to Assign Licenses"

        If ($User) {
            foreach ($Assignment in $LicenseSelection) {

                $addLicenses = @(
                    @{
                        SkuId         = $Assignment.LicenseId
                        DisabledPlans = $Assignment.DisabledServicePlans
                    }
                )

                foreach ($User in $Users) {

                    Set-MgUserLicense -UserId $user.Id -AddLicenses $addLicenses -RemoveLicenses @()
                
                }
            }
        }
        else {
            Write-Host "No User Selected. Exiting" -ForegroundColor Yellow
            Exit
        }
    }elseif($UserOrGroup -eq "G") {
        [array]$Groups = Get-MgGroup -All -Filter "securityEnabled eq true"  | Out-GridView -PassThru -Title "Select Groups to Assign Licenses"

        If ($Groups) {
            foreach ($Assignment in $LicenseSelection) {

                $addLicenses = @(
                    @{
                        SkuId         = $Assignment.LicenseId
                        DisabledPlans = $Assignment.DisabledServicePlans
                    }
                )

                foreach ($Group in $Groups) {

                    Set-MgGroupLicense -GroupId $Group.Id -AddLicenses $addLicenses -RemoveLicenses @()
                
                }
            }
        }
        else {
            Write-Host "No Group Selected. Exiting" -ForegroundColor Yellow
            Exit
        }
    }
}
else {
    write-host "No Licenses Selected. Exiting" -ForegroundColor Yellow
    Exit
}

