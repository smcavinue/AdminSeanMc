##Error Handling PowerShell Examples##

##An example of a cmdlet that will give you an error
$Process = Get-Process -Name "NoProcessWithThisNameExists"
Write-Host "The process ID for $($Process.Name) is $($Process.Id)"

##Update the ErrorActionPreference to Stop
$ErrorActionPreference = "Stop"

##An example of a cmdlet that will stop the script if an error occurs
$Process = Get-Process -Name "NoProcessWithThisNameExists"##This line will not be displayed and the script will stop

##Reset the ErrorActionPreference to Continue
$ErrorActionPreference = "Continue"

##An example of a cmdlet that will stop the script if an error occurs
Get-Process -Name "NoProcessWithThisNameExists" -ErrorAction Stop
Write-Host "The process ID for $($Process.Name) is $($Process.Id)"##When run in a script, this line will not be displayed and the script will stop

##Update the ErrorActionPreference to Stop
$ErrorActionPreference = "Stop"
Get-Process -Name "NoProcessWithThisNameExists" -ErrorAction silentlycontinue
Write-Host "I don't care if that last cmdlet failed"
Get-Process -Name "NoProcessWithThisNameExists" 
Write-Host "The process ID for $($Process.Name) is $($Process.Id)"##When run in a script, this line will not be displayed and the script will stop

##Using a Try/Catch block for manage an error
Try {
    $user = Get-MgUser -UserId Userdoesntexist@seanmcavinue.net -ErrorAction Stop
    Write-host "The users name is: $($user.DisplayName)"
}catch{
    Write-Host "User doesn't exist!"
}

##Adding in exception details in the output
Try {
    $user = Get-MgUser -UserId Userdoesntexist@seanmcavinue.net -ErrorAction Stop
    Write-host "The users name is: $($user.DisplayName)"
}catch{
    Write-Host "User doesn't exist!"
    Write-Host "By the way, the Error we got was: $($_.Exception.Message)"
}

##Adding a Finally block
Try {
    $user = Get-MgUser -UserId Userdoesntexist@seanmcavinue.net -ErrorAction Stop
    Write-host "The users name is: $($user.DisplayName)"
}catch{
    Write-Host "User doesn't exist!"
    Write-Host "By the way, the Error we got was: $($_.Exception.Message)"
}Finally{
    Disconnect-MgGraph
    Write-Host "I don't really care if the last stuff worked, but I've logged you out!"
}