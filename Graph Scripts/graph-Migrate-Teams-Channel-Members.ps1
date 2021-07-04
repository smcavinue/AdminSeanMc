<##Author: Sean McAvinue
##Details: Graph / PowerShell Script to migrate Teams channels and and channel messages between tenants, 
##          Please fully read and test any scripts before running in your production environment!
        .SYNOPSIS
        Copies Teams membership from one Team to another. Also copies membership of any private channels that have the same name in both Teams

        .PARAMETER SourceTeamObjectID
        Required - The GroupID of the Team in the source

        .PARAMETER SourceClientID
        Required - Application (Client) ID of the App Registration

        .PARAMETER SourceTenantID
        Required - Directory (Tenant) ID of the Azure AD Tenant

        .PARAMETER SourceClientSecret
        Required - Client Secret for the Source Application Registration

        .PARAMETER TargetClientID
        Required - Application (Client) ID of the App Registration

        .PARAMETER TargetTenantID
        Required - Directory (Tenant) ID of the Azure AD Tenant

        .PARAMETER Mappingfile
        Required - Path to Mapping file CSV

        .PARAMETER TargetTeamObjectID
        Required - the GroupID of the Team in the target, if blank, a new Team will be created

        .PARAMETER TargetToken
        Required - A Delegated Graph Token for the Target Tenancy. Can be optained from MSAL.PS - Get-MSALToken



        .EXAMPLE
        This example will copy membership from the Team with ID 2b17feef-f7bc-4928-a351-472dfb7cd115 in the source tenant to the Team with ID 3991e736-5c6d-4eb2-9c70-5f60c509833d in the Target Tenant
        .\graph-Migrate-Teams-Channel-Members.ps1 -SourceTeamObjectID fa9377ca-0824-4da5-aea9-496856298ecc -SourceClientSecret $SourceclientSecret -SourceClientID $SourceclientID -SourceTenantID $SourcetenantID -TargetClientID $TargetclientID -TargetTenantID $TargettenantID -TargetTeamObjectID 49736ecc-0755-4ec8-abba-6995e7c34a05 -MappingFile C:\temp\Mappingfile.csv -TargetToken $TargetToken.accesstoken
        .Notes
        For similar scripts check out the links below
        
            Blog: https://seanmcavinue.net
            GitHub: https://github.com/smcavinue
            Twitter: @Sean_McAvinue
            Linkedin: https://www.linkedin.com/in/sean-mcavinue-4a058874/


    #>

##

Param(
    [parameter(Mandatory = $true)]
    [String]
    $SourceTeamObjectID,
    [parameter(Mandatory = $true)]
    [String]
    $SourceClientID,
    [parameter(Mandatory = $true)]
    [String]
    $SourceTenantID,
    [parameter(Mandatory = $true)]
    [String]
    $TargetClientID,
    [parameter(Mandatory = $true)]
    [String]
    $TargetTenantID,
    [parameter(Mandatory = $true)]
    [String]
    $Mappingfile,
    [parameter(Mandatory = $true)]
    [String]
    $SourceClientSecret,
    [parameter(Mandatory = $true)]
    [String]
    $TargetTeamObjectID,
    [parameter(Mandatory = $true)]
    [String]
    $TargetToken
)

##FUNCTIONS##
function GetGraphToken {
    # Azure AD OAuth Application Token for Graph API
    # Get OAuth token for a AAD Application (returned as $token)
    <#
        .SYNOPSIS
        This function gets and returns a Graph Token using the provided details
    
        .PARAMETER clientSecret
        -is the app registration client secret
    
        .PARAMETER clientID
        -is the app clientID
    
        .PARAMETER tenantID
        -is the directory ID of the tenancy
        
        #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $ClientSecret,
        [parameter(Mandatory = $true)]
        [String]
        $ClientID,
        [parameter(Mandatory = $true)]
        [String]
        $TenantID
    
    )
    
   
        
    # Construct URI
    $uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
         
    # Construct Body
    $body = @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }
         
    # Get OAuth 2.0 Token
    $tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
         
    # Access Token
    $token = ($tokenRequest.Content | ConvertFrom-Json).access_token
    return $token
}

function RunQueryandEnumerateResults {
    <#
    .SYNOPSIS
    Runs Graph Query and if there are any additional pages, parses them and appends to a single variable
    
    .PARAMETER apiUri
    -APIURi is the apiUri to be passed
    
    .PARAMETER token
    -token is the auth token
    
    #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $apiUri,
        [parameter(Mandatory = $true)]
        $token

    )
    ##write-host "DEBUG: Running $apiuri"
    #Run Graph Query
    $Results = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)
    #Output Results for debug checking
    #write-host $results

    #Begin populating results
    $ResultsValue = $Results.value

    #If there is a next page, query the next page until there are no more pages and append results to existing set
    if ($results."@odata.nextLink" -ne $null) {
        #write-host "enumerating pages..." -ForegroundColor yellow
        start-sleep -seconds 2
        $NextPageUri = $results."@odata.nextLink"
        ##While there is a next page, query it and loop, append results
        While ($NextPageUri -ne $null) {
            $NextPageRequest = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $NextPageURI -Method Get)
            $NxtPageData = $NextPageRequest.Value
            $NextPageUri = $NextPageRequest."@odata.nextLink"
            $ResultsValue = $ResultsValue + $NxtPageData
        }
    }

    ##Return completed results
    return $ResultsValue

    
}

function WriteLog {
    <#
    .SYNOPSIS
    Creates temp directory at C:\temp and logs to a specified file in this folder
    
    .PARAMETER LogEntry
    String to be added to the log

    .PARAMETER LogFileName
    Name of the log file to write to in C:\Temp
    

    #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $LogEntry,
        [parameter(Mandatory = $true)]
        [String]
        $LogFileName
    )
    
    ##Get Timestamp for log
    [String]$DateTime = get-date

    ##Create Temp folder if it's missing
    if (!(test-path "c:\temp")) {
        New-Item -Path "c:\" -Name "temp" -ItemType "directory"  
    }

    ##Construct full log entry with timestamp prefix
    $logentry = "$DateTime : $LogEntry"
    write-host $LogEntry -ForegroundColor green
    ##Write to file
    $LogEntry | out-file "c:\temp\$LogFileName" -Append
}

function GetTeamChannels {
    <#
    .SYNOPSIS
    Returns the Team Channels for the given Team ID
    
    .PARAMETER Token
    Graph Access Token

    .PARAMETER TeamID
    Group ID of the Team
    

    #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $Token,
        [parameter(Mandatory = $true)]
        [String]
        $TeamID
    )

    $apiUri = "https://graph.microsoft.com/beta/teams/$TeamID/channels"
    
    $channels = RunQueryandEnumerateResults -apiUri $apiUri -token $token

    Return $channels
}

function GetTeamMembers {
    <#
        .SYNOPSIS
        Returns the Team members for the given Team ID
        
        .PARAMETER Token
        Graph Access Token
    
        .PARAMETER TeamID
        Group ID of the Team
        
    
        #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $Token,
        [parameter(Mandatory = $true)]
        [String]
        $TeamID
    )
    
    $apiUri = "https://graph.microsoft.com/beta/teams/$TeamID/members"
    $members = RunQueryandEnumerateResults -apiUri $apiUri -token $sourcetoken
    
    Return $members
}
    
function GetTeam {
    <#
    .SYNOPSIS
    Returns the Team Object for the given Team ID
    
    .PARAMETER Token
    Graph Access Token

    .PARAMETER TeamID
    Group ID of the Team
    

    #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $Token,
        [parameter(Mandatory = $true)]
        [String]
        $TeamID
    )

    $apiUri = "https://graph.microsoft.com/beta/teams/$TeamID"
    $Team = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)" } -Uri $apiUri -Method Get)

    Return $Team
}


function AddTeamMembers {
    <#
    .SYNOPSIS
    Short description
    
    .PARAMETER Token
    Graph Access Token
    
    .PARAMETER TeamID
    Target Team ID
    
    .PARAMETER SourceMembers
    Source Members Object
    
    .PARAMETER MappingFile
    Path to CSV Mapping File

    #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $Token,
        [parameter(Mandatory = $true)]
        [String]
        $TeamID,
        [parameter(Mandatory = $true)]
        [Object]
        $SourceMembers,
        [parameter(Mandatory = $true)]
        [String]
        $MappingFile
    )

    $Mappings = import-csv $MappingFile

    foreach ($member in $SourceMembers) {

        $TargetID = ($mappings | where-object { $_.sourceid -eq $member.userid } | Select-Object -First 1).targetid

        if ($null -ne $targetID) {
            if ($member.roles -eq "guest") {
                $body = @"
        {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": ["guest"],
            "user@odata.bind": "https://graph.microsoft.com/v1.0/users('$targetID')"
        }
"@
            }
            elseif ($member.roles -eq "owner") {
                $body = @"
        {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": ["owner"],
            "user@odata.bind": "https://graph.microsoft.com/v1.0/users('$targetID')"
        }
"@
            }
            else {
                $body = @"
        {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "user@odata.bind": "https://graph.microsoft.com/v1.0/users('$targetID')"
        }
"@

            }

            $apiuri = "https://graph.microsoft.com/v1.0/teams/$teamID/members"
            writelog -LogEntry "Adding $targetID" -LogFileName teams-migrationlog.log

            Try {
                $memberAdded = Invoke-WebRequest -Method Post -Uri $apiuri -ContentType "application/json" -Body $body -UseBasicParsing -Headers @{Authorization = "Bearer $($Token)" } -ErrorAction Continue
                start-sleep -Seconds 1
                writelog -LogEntry "Added $targetID successfully" -LogFileName teams-migrationlog.log
            }
            catch {
                write-host "$Message $($_.Exception.Message)"
                writelog -LogEntry "Error Adding $targetID" -LogFileName teams-migrationlog.log
            }
            start-sleep -Seconds 1
        }
    }
        
    writelog -LogEntry "Waiting for membership replication" -LogFileName teams-migrationlog.log
    start-sleep -Seconds 10
}



function ProcessChannels {
    <#
    .SYNOPSIS
    Processes channels from the source team and provisions any missing channels in the target
    
    .PARAMETER SourceToken
    Source Access Token
    
    .PARAMETER TargetToken
    Target Access Token
    
    .PARAMETER SourceChannels
    Source Channels Object
    
    .PARAMETER TargetTeam
    Target Team ID
    
    .PARAMETER SourceTeam
    Source Team ID

    .PARAMETER MappingFile
    
    #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $SourceToken,
        [parameter(Mandatory = $true)]
        [String]
        $TargetToken,
        [parameter(Mandatory = $true)]
        [Object]
        $SourceChannels,
        [parameter(Mandatory = $true)]
        [String]
        $TargetTeam,
        [parameter(Mandatory = $true)]
        [String]
        $SourceTeam,
        [parameter(Mandatory = $true)]
        [String]
        $MappingFile
    )
      
    $TargetChannels = GetTeamChannels -token $TargetToken -TeamID $TargetTeam

    foreach ($channel in $SourceChannels) {
        $existingChannel = ($TargetChannels | Where-Object { $_.displayname -eq $channel.displayname })

        if ($existingChannel) {
            writelog -LogEntry "Channel $($channel.displayname) exists... adding membership" -LogFileName "teams-migrationlog.log"
        }
        else {
            writelog -LogEntry "Channel $($channel.displayname) doesn't exist or is public... will be skipped for membership" -LogFileName "teams-migrationlog.log"

        }

        if (($Channel.membershipType -eq "private") -and ($existingChannel.id)) {
            writelog -LogEntry "Assigning private channel membership for $($existingChannel.displayname)" -LogFileName "teams-migrationlog.log"
            
            AssignPrivateChannelMembership -sourceToken $SourceToken -targetToken $TargetToken -SourceTeam $SourceTeam -SourceChannel $Channel.id -TargetTeam $TargetTeam -TargetChannel $existingChannel.id  -mappingfile $Mappingfile

        }

    }
}


function AssignPrivateChannelMembership {
    <#
    .SYNOPSIS
    Assigns membership to Private channels
    
    .PARAMETER SourceToken
    Source Access Token
    
    .PARAMETER TargetToken
    Target Access Token
    
    .PARAMETER SourceChannel
    Source Channel ID
    
    .PARAMETER TargetTeam
    Target Team ID
    
    .PARAMETER SourceTeam
    Source Team ID

    .PARAMETER TargetChannel
    Source Channel ID

    .PARAMETER MappingFile
    Path to Mapping File
    
    #>
    Param(
        [parameter(Mandatory = $true)]
        [String]
        $SourceToken,
        [parameter(Mandatory = $true)]
        [String]
        $TargetToken,
        [parameter(Mandatory = $true)]
        [String]
        $SourceChannel,
        [parameter(Mandatory = $true)]
        [String]
        $TargetChannel,
        [parameter(Mandatory = $true)]
        [String]
        $SourceTeam,
        [parameter(Mandatory = $true)]
        [String]
        $TargetTeam,
        [parameter(Mandatory = $true)]
        [String]
        $Mappingfile
    )

    $apiuri = "https://graph.microsoft.com/beta/teams/$($SourceTeam)/channels/$($SourceChannel)/members"

    #write-host $apiuri -foregroundcolor yellow
    $SourceMembers = RunQueryandEnumerateResults -token $sourceToken -apiuri $apiuri

    $Mappings = import-csv $Mappingfile

    foreach ($member in $sourcemembers) {
        $TargetMember = $null
        $TargetMember = (($mappings | where-object { $_.sourceid -eq $member.userid } | Select-Object -First 1).targetid)

        if ($TargetMember -like "*-*") {
            $apiuri = "https://graph.microsoft.com/beta/teams/$($TargetTeam)/channels/$($TargetChannel)/members"


            if ($member.roles -eq "guest") {
                $User = @"
        {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": ["$($Member.roles)"],
            "user@odata.bind": "https://graph.microsoft.com/v1.0/users('$TargetMember')"
        }
"@
            }
            elseif ($member.roles -eq "Owner") {
                $User = @"
        {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": ["$($Member.roles)"],
            "user@odata.bind": "https://graph.microsoft.com/v1.0/users('$TargetMember')"
        }
"@
            }
            else {
                $User = @"
        {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "user@odata.bind": "https://graph.microsoft.com/v1.0/users('$TargetMember')"
        }
"@
            }

            writelog -LogEntry "Adding $($member.displayname) to Channel $Targetchannel" -LogFileName "teams-migrationlog.log"

            Try {
                $MemberAdded = Invoke-WebRequest -Method Post -Uri $apiuri -ContentType "application/json" -Body $User -UseBasicParsing -Headers @{Authorization = "Bearer $($TargetToken)" } -ErrorAction stop
                start-sleep -Seconds 1
                writelog -LogEntry "Added $($member.displayname) successfully" -LogFileName "teams-migrationlog.log"
            }
            catch {
                writelog -LogEntry "Error Adding $($member.displayname)" -LogFileName "teams-migrationlog.log"
            }
        }
        else {
            writelog -LogEntry  "Skipping user $($member.userid) not in mapping file"  -LogFileName "teams-migrationlog.log"
        }
    }

}



##################################################################################################################Main###################################################################################################
##Initiate Log
$LogFileName = "teams-migrationlog.log"
try {
    WriteLog -LogEntry "Migration starting for $SourceTeamObjectID" -LogFileName $LogFileName
}
catch {
    $Message = "Error Creating Log file, please check permissions and mapping file"
    throw "$Message $($_.Exception.Message)"
    break
}

##Try to retrieve Token for Source Tenant
try {
    $Sourcetoken = GetGraphToken -ClientSecret $SourceClientSecret -ClientID $SourceClientID -TenantID $SourceTenantID -ErrorAction stop
    

}
catch {
    $Message = "Error Getting Source Token, please check Source Application Registration Details"
    WriteLog -LogEntry $Message -LogFileName $LogFileName
    throw "$Message $($_.Exception.Message)"
    break
}

WriteLog -LogEntry "Source Token Retrieved" -LogFileName $LogFileName

##Get Source Team Object
try {
    $SourceTeam = GetTeam -Token $Sourcetoken -TeamID $SourceTeamObjectID
}
catch {

    $Message = "Error Getting Source Team, please check Source Team ID provided is correct"
    WriteLog -LogEntry $Message -LogFileName $LogFileName
    throw "$Message $($_.Exception.Message)"
    break

}
WriteLog -LogEntry "Source Team Retrieved" -LogFileName $LogFileName

##Get Source Team Channels
try {
    $SourceChannels = GetTeamChannels -Token $Sourcetoken -TeamID $SourceTeamObjectID
}
catch {
    
    $Message = "Error Getting Source Team channels, please check API Permissions"
    WriteLog -LogEntry $Message -LogFileName $LogFileName
    throw "$Message $($_.Exception.Message)"
    break
    
}
WriteLog -LogEntry "Source Team channels Retrieved" -LogFileName $LogFileName

##Get Source Team Members
try {
    $SourceMembers = GetTeamMembers -Token $Sourcetoken -TeamID $SourceTeamObjectID
}
catch {
    
    $Message = "Error Getting Source Team Members, please check API Permissions"
    WriteLog -LogEntry $Message -LogFileName $LogFileName
    throw "$Message $($_.Exception.Message)"
    break
    
}
WriteLog -LogEntry "Source Team members Retrieved" -LogFileName $LogFileName


##Add Members to target Team
AddTeamMembers -Token $Targettoken -TeamID $TargetTeamObjectID -SourceMembers $SourceMembers -mappingfile $Mappingfile

##Create Channels in the target team if they don't exist
try {
    ProcessChannels -SourceToken $Sourcetoken -TargetToken $Targettoken -SourceChannels $SourceChannels -TargetTeam $TargetTeamObjectID -SourceTeam $SourceTeamObjectID -mappingfile $Mappingfile
}
catch {
    
    $Message = "Error Updating Channels, please check API Permissions"
    WriteLog -LogEntry $Message -LogFileName $LogFileName
    throw "$Message $($_.Exception.Message)"
    break
    
}


