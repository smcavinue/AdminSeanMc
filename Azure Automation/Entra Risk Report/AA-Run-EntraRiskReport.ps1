<#
    .SYNOPSIS
        This script will connect to the Microsoft Graph Security API and retrieve any risk detections, risky users and risky service principals.
        It will then format the results into an HTML report and send it via email. This is the Azure Automation version of the script that can be run in Azure Automation.

    .DESCRIPTION
        This script will connect to the Microsoft Graph Security API and retrieve any risk detections, risky users and risky service principals.
        It will then format the results into an HTML report and send it via email.

    .EXAMPLE
        Run the script with the default parameters.
        .\AA-Run-EntraRiskReport.ps1

    .NOTES
        NAME:    AA-Run-EntraRiskReport.ps1

#>

##Define email values
$sender = "sean.mcavinue@seanmcavinue.net"
$recipient = "sean.mcavinue@seanmcavinue.net"
$subject = "Daily Risk Report"

##Connect to Microsoft Graph
Connect-MgGraph -Identity -NoWelcome

[array]$RiskDetections = Get-MgRiskDetection -All -filter "riskstate eq 'atRisk'"
[array]$RiskyUsers = Get-MgRiskyUser -All -filter "riskstate eq 'atRisk'"
[array]$RiskyServicePrincipal = Get-MgRiskyServicePrincipal -All -filter "riskstate eq 'atRisk'"

##Add each risk detection to a collection
$RiskDetectionCollection = @()
foreach ($RiskDetection in $RiskDetections) {
    $RiskDetectionCollection += [PSCustomObject]@{
        "Id" = $RiskDetection.id
        "RiskDetectionDateTime" = $RiskDetection.DetectedDateTime
        "RiskLevel" = $RiskDetection.riskLevel
        "RiskState" = $RiskDetection.riskState
        "RiskDetail" = $RiskDetection.riskDetail
        "RiskType" = $RiskDetection.riskEventType
        "UserId" = $RiskDetection.userId
        "UserDisplayName" = $RiskDetection.userDisplayName
        "UserPrincipalName" = $RiskDetection.userPrincipalName
        "RiskLastUpdatedDateTime" = $RiskDetection.LastUpdatedDateTime
        "RiskEventTypes" = $RiskDetection.riskEventType
    }
}

##Add each risky user to a collection
$RiskyUserCollection = @()
foreach ($RiskyUser in $RiskyUsers) {
    $RiskyUserCollection += [PSCustomObject]@{
        "Id" = $RiskyUser.id
        "RiskLastUpdatedDateTime" = $RiskyUser.RiskLastUpdatedDateTime
        "RiskLevel" = $RiskyUser.riskLevel
        "RiskState" = $RiskyUser.riskState
        "RiskDetail" = $RiskyUser.riskDetail
        "UserDisplayName" = $RiskyUser.userDisplayName
        "UserPrincipalName" = $RiskyUser.userPrincipalName
    }
}

##Add each risky service principal to a collection
$RiskyServicePrincipalCollection = @()
foreach ($RiskyServicePrincipal in $RiskyServicePrincipal) {
    $RiskyServicePrincipalCollection += [PSCustomObject]@{
        "Id" = $RiskyServicePrincipal.id
        "AppID" = $RiskyServicePrincipal.appId
        "ServicePrincipalDisplayName" = $RiskyServicePrincipal.DisplayName
        "RiskLastUpdatedDateTime" = $RiskyServicePrincipal.RiskLastUpdatedDateTime
        "RiskLevel" = $RiskyServicePrincipal.riskLevel
        "RiskState" = $RiskyServicePrincipal.riskState
        "ServicePrincipalType" = $RiskyServicePrincipal.servicePrincipalType
    }
}

##Export all collections to a single HTML report
$HTML = ($RiskDetectionCollection | ConvertTo-Html -Head $htmlHead -Body $htmlBody -Title "Risk Detections" -PreContent "<h1>Risk Detections</h1>")
$HTML += ($RiskyUserCollection | ConvertTo-Html -Head $htmlHead -Body $htmlBody -Title "Risky Users" -PreContent "<h1>Risky Users</h1>")
$HTML += ($RiskyServicePrincipalCollection | ConvertTo-Html -Head $htmlHead -Body $htmlBody -Title "Risky Service Principals" -PreContent "<h1>Risky Service Principals</h1>")

# Add CSS styles to format the tables
$CSS = @"
<style>
    table {
        border-collapse: collapse;
        font-size: 10px;
        width: 100%;
    }
    th, td {
        border: 1px solid black;
        padding: 8px;
        text-align: left;
    }
    th {
        background-color: #f2f2f2;
    }
    .low {
        background-color: #00FF00;
    }
    .medium {
        background-color: #FFFF00;
    }
    .high {
        background-color: #FF0000;
    }
</style>
"@

# Insert the CSS styles into the HTML file
$HTML = $HTML -replace "<head>", "<head>`n$CSS`n"
$HTML = $HTML -replace "<td>low</td>", "<td style=`"background-color: #00FF00;`">low</td>"
$HTML = $HTML -replace "<td>medium</td>", "<td style=`"background-color: #FFFF00;`">medium</td>"
$HTML = $HTML -replace "<td>high</td>", "<td style=`"background-color: #FF0000;`">high</td>"

##Send the report via email
$body = $HTML.ToString()

$params = @{
    Message         = @{
        Subject       = $subject
        Body          = @{
            ContentType = "HTML"
            Content     = "$($HTML)"
        }
        ToRecipients  = @(
            @{
                EmailAddress = @{
                    Address = $recipient
                }
            }
        )
    }
    SaveToSentItems = $save
}
# Send message
Send-MgUserMail -UserId $sender -BodyParameter $params