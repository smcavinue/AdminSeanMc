##Report on Teams and Groups with Guest Users##
#This script will report on all Teams and Groups in your tenant that have Guest Users. It will output the results to a CSV file.
#The script will also output the results to the console at runtime
#The script will prompt you to login to your tenant and consent to the required permissions.
#The Microsoft.Graph PowerShell Module is requires to run this script
#You can install the Microsoft.Graph module by running Install-Module -Name Microsoft.Graph
#Change the below line to specify the output CSV Path
$Outputfile = "c:\temp\GuestReport.csv"

#Connect to Microsoft Graph
Try{
    Connect-MgGraph -Scopes "Group.Read.All","User.Read.All" -NoWelcome
    Write-Host "Connected to the Microsoft Graph"
}
Catch{
    Write-Host "Error connecting to Microsoft Graph. Please try again."
    Write-Host $_.Exception.Message
    Exit
}

Try{
    #Get all Groups
    [array]$Groups = Get-MgGroup -All
    Write-Host "Groups found in Tenant: $($Teams.Count)"
}
Catch{
    Write-Host "Error getting Groups. Please try again."
    Write-Host $_.Exception.Message
    Exit
}

$Results = @()
$x = 0
foreach($Group in $Groups){
    $x++
    Write-Progress -Activity "Checking Groups for Guest Users" -Status "Checking Group $($x) of $($Groups.Count)" -PercentComplete (($x / $Groups.Count) * 100)
    #Get all Members of the Group
    Try{
        [array]$Members = Get-MgGroupMember -GroupId $Group.Id
    }
    Catch{
        Write-Host "Error getting Members of Group $($Group.DisplayName). Please try again."
        Write-Host $_.Exception.Message
        Continue
    }

    #Check if any Members are Guests
    $Guests = $Members | Where-Object {$_.AdditionalProperties.userPrincipalName -like "*#EXT#*"}

    if($Guests.Count -gt 0){
        Write-Host "Group $($Group.DisplayName) has $($Guests.Count) Guest Users" -ForegroundColor Yellow
        $Output = [PSCustomObject]@{
            GroupName = $Group.DisplayName
            GroupId = $Group.Id
            GuestCount = $Guests.Count
            GuestList = $Guests.AdditionalProperties.userPrincipalName -join ";"
        }
        $Results += $Output
    }else{
        Write-Host "Group $($Group.DisplayName) has $($Members.count) members but no Guest Users"
    }
}
$Results | Export-Csv -Path $Outputfile

Write-Host "Report Complete. Results saved to $Outputfile" -ForegroundColor Green
