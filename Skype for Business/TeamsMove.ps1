##This function allows us to select users who are still homed in Skype for business
function move-toTeams-selectusers{
    $i = 1
    ##Allow selection of one or multiple users
    $users = (Get-CsUser -ResultSize unlimited | ?{$_.registrarPool -notlike $null} | Out-GridView -PassThru -Title "Select One or more users to migrate")
    
    ##Writing Progress to Screen
    Write-host you selected $users.count users

    ##check if it's array
    if($users -is [array]){
        foreach($User in $users){
        
             write-host Moving user $i of $users.count : $user.SipAddress

             Move-ToTeams-performMove $user
             $i++
        }
    }else{
            
            write-host Moving single user: $users.SipAddress
            Move-ToTeams-performMove $users

     }
}

function Move-ToTeams-performMove{
    <#
    .SYNOPSIS
    This function performs the move from Skype on Premises to Teams
    
    .DESCRIPTION
    This function accepts a user account and moves the user to Teams
    
    .PARAMETER useridentity
    Takes a user in for processing by the migration process
        
    .NOTES
    General notes
    #>
    param($useridentity)

    ##Removes conferencing policy from user account
     Grant-CsConferencingPolicy -PolicyName "No Dial-in" -Identity $useridentity.identity
     
     ##Create Dialog
     $a = new-object -comobject wscript.shell 

     ##Prompt for if calling policy should be enabled
     $intAnswer = $a.popup("Should " + $useridentity.sipaddress + " be enabled for outbound calling", 0,"Outbound Calls",4) 
    
    ##IF yes
    If ($intAnswer -eq 6) { 
        ##warn admin to ensure license is assigned
       $a.popup("Enabling Calling Policy for " + $useridentity.sipaddress + ", make sure they are licensed for Calling or this will fail!") 
        ##Enable outbound calling
        Set-CsUser -Identity $useridentity.identity -EnterpriseVoiceEnabled:$True
    }##Else no 
    else { 
        $a.popup("Removing Calling Policy from " + $useridentity.sipaddress + ", license can be removed if assigned") 
        ##Disable outbound calling
        Set-CsUser -Identity $useridentity.identity -EnterpriseVoiceEnabled:$False -Confirm:$false
    } 
  
    Try{
        ##Moves user to Teams
        Move-CsUser -Identity $UserIdentity.identity -Target sipfed.online.lync.com -MoveToTeams -Credential $credentials -HostedMigrationOverrideUrl $url -ErrorAction Stop -Confirm:$false
        write-host $UserIdentity.identity completed successfully -ForegroundColor Green

    }catch{
    #Catch and notify admin of error
    $ErrorMessage = $_.Exception.Message

    write-host Failed Migrating $useridentity.sip with error $ErrorMessage

    }

}


function Move-toTeams{



    ##Adding Teams endpoint URL
    $url="https://admin0e.online.lync.com/HostedMigration/hostedmigrationService.svc"
    
    ##If credentials dont exist, prompt for them
    if(!($credentials)){

        $credentials = Get-Credential -Message "Please enter credentials for an Office 365 admin with onmicrosoft UPN"

    }

    move-toTeams-selectusers

}