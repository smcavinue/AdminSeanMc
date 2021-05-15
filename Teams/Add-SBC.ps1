##Autor: Sean McAvinue
##Web: https://seanmcavinue.net
##Twitter: @Sean_McAvinue
##GitHub: https://github.com/smcavinue/AdminSeanMc
Function AddSBC {
    <#
    .SYNOPSIS
    Adds and SBC for Teams Direct Routing
    
    .PARAMETER PSTNUsageName
    -A name for the new PSTN usage record
    
    .PARAMETER FQDN
    -FQDN of the SBC

    .PARAMETER SBCDescription
    -Description for the SBC

    .PARAMETER SIPSignalingPort (Optional)
    -the SIP Signaling Port, default value is 5067

    .PARAMETER SendSIPOptions (Optional)
    -Option to Send SIP Options, Default is $TRUE

    .PARAMETER ForwardPai (Optional)
    -Option to forward Pai Header, Default is $TRUE

    .PARAMETER FailoverResponseCodes (Optional)
    -the supportd failover response codes for direct routing, default is '408, 503, 504'

    .PARAMETER FailoverTimeSeconds (Optional)
    -the failover time (seconds), default is 10

    .PARAMETER PidfLoSupported (Optional)
    -SBC supports PIDF/LO for emergency calls, default is $FALSE

    .PARAMETER MaxConcurrentSessions (Optional)
    -Max supported concurrent sessions by SBC

    .PARAMETER VoiceRoutingPolicyName
    -Name for the Voice Routing Policy

    .PARAMETER VoiceRouteName
    -Name for the Voice Route

    .PARAMETER NumberPattern (Optional)
    -Number Pattern for the Voice Route

    .PARAMETER ForwardCallHistory (Optional)
    -Should call history be forwarded, default is $FALSE

    .Example
    AddSBC -PSTNUsageName "SBC-EU" -FQDN "SBCEU.adminseanmc.com" -SBCDescription "EU Primary SBC" -SIPSignalingPort 5067 -SendSIPOptions $true -ForwardPai $true -FailoverResponseCodes '408, 503, 504' -FailoverTimeSeconds "10" -PidfLoSupported $false -MaxConcurrentSessions "200" -VoiceRoutingPolicyName "SBC-EU-VRP" -VoiceRouteName "SBC-EU-VR" -NumberPattern "^(\+[0-9]{7,15})$" -ForwardCallHistory $false
    #>

    Param(
        [parameter(Mandatory = $true)]
        $PSTNUsageName,
        [parameter(Mandatory = $true)]
        $FQDN,
        [parameter(Mandatory = $true)]
        $SBCDescription,
        [parameter(Mandatory = $false)]
        $SIPSignalingPort = "5067",
        [parameter(Mandatory = $false)]
        $SendSIPOptions = $True,
        [parameter(Mandatory = $false)]
        $ForwardPai = $True,
        [parameter(Mandatory = $false)]
        $FailoverResponseCodes = '408, 503, 504',
        [parameter(Mandatory = $false)]
        $FailoverTimeSeconds = "10",
        [parameter(Mandatory = $false)]
        $PidfLoSupported = $false,
        [parameter(Mandatory = $false)]
        $MaxConcurrentSessions = 20,
        [parameter(Mandatory = $true)]
        $VoiceRoutingPolicyName,
        [parameter(Mandatory = $true)]
        $VoiceRouteName,
        [parameter(Mandatory = $false)]
        $NumberPattern = "^(\+[0-9]{7,15})$",
        [parameter(Mandatory = $false)]
        $ForwardCallHistory = $False
        
    )

    ##Adds PSTN Usage Record##
    $CurrentPSTNUsage = Get-CsOnlinePstnUsage

    if (!($CurrentPSTNUsage.usage -contains "SBC")) {
        
        write-host "$PSTNUsageName does not exist, creating new" -ForegroundColor yellow
        Set-CsOnlinePstnUsage -Identity global -Usage @{add = "$PSTNUsageName" }

    }
    else {
    
        write-host "$PSTNUsageName exists, using existing" -ForegroundColor yellow

    }

    write-host "Pausing for PSTN Usage Replication" -ForegroundColor yellow
    start-sleep 100

    ##Add SBC to Direct Routing Configuration##
    write-host "Creating SBC"
    try {

        New-CsOnlinePSTNGateway -Fqdn $FQDN -Description $SBCDescription -Enabled $true -SipSignalingPort $SIPSignalingPort -SendSipOptions $SendSIPOptions -ForwardCallHistory $ForwardCallHistory -ForwardPai $ForwardPai -FailoverResponseCodes $FailoverResponseCodes -FailoverTimeSeconds $FailoverTimeSeconds -PidfLoSupported $PidfLoSupported -MaxConcurrentSessions $MaxConcurrentSessions -ErrorAction stop                   
    
    }
    catch {

        Write-host "Error Creating SBC, stopping!" $Error -ForegroundColor red

    }

    write-host "Pausing for SBC Replication" -ForegroundColor yellow
    start-sleep 100

    ##Add Voice Routing Policy##
    if (Get-CsOnlineVoiceRoutingPolicy -Identity $VoiceRoutingPolicyName -ErrorAction silentlycontinue) {
        write-host "Voice Routing Policy $VoiceRoutingPolicyName  exists already, using existing" -ForegroundColor yellow
    }
    else {

        write-host "Creating Voice Routing Policy $VoiceRoutingPolicyName" -ForegroundColor yellow
        New-CsOnlineVoiceRoutingPolicy -Identity $VoiceRoutingPolicyName -OnlinePstnUsages $PSTNUsageName

    }
    write-host "Pausing for Voice Routing Policy Replication" -ForegroundColor yellow
    start-sleep 100

    ##Create a new Voice Route
    if (get-csonlinevoiceroute -Identity $VoiceRouteName -ErrorAction silentlycontinue) {

        write-host "Voice Route $VoiceRouteName already exists, updating existing"
        set-CsOnlineVoiceRoute -Identity $VoiceRouteName -OnlinePstnUsages @{add = "$PSTNUsageName" } -OnlinePstnGatewayList @{add = "$FQDN" } 

    }
    else {

        New-CsOnlineVoiceRoute -Identity $VoiceRouteName -OnlinePstnUsages @{add = "$PSTNUsageName" } -OnlinePstnGatewayList @{add = "$FQDN" } -NumberPattern $NumberPattern

    }


}

