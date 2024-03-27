# ImportMSTeamsCallQueue.ps1
# Use CSV file to create Call Queues.
#
# REQUIREMENTS:
# PS-Module Microsoft-Teams and Microsoft.Graph (Optional)
# Graph-Api authorizations needed: User.Read.All, Group.Read.All, TeamSettings.Read.All
# Export Worksheet Call Queues from MSTeams-Runbook.xlsx as CSV file
# CreateMSTeamsResourceAccounts.ps1 should be run before.
#
# Version 1.0.1 (Build 1.0.1-2024-03-27)
# 
# Created by: Armin Toepper
#
#########################################################################################
#
#
#########################################################################################
#                            DO NOT EDIT BELOW THESE LINES!                             #
#########################################################################################
#
Set-StrictMode -Version "2.0"
$ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
Add-Type -AssemblyName System.Windows.Forms

function GetBool {
    Param (
        [string]$onoff
    )

    if ($onoff -eq "On") {
        return $true
    } else {
        return $false
    }
}

Clear-Host
$useOfMG = Read-Host -Prompt "Do you want to use GraphsApi? [y/N]"

Clear-Host
Write-Host "Please enter your Tenant ID (can be left empty if you have only one)" -ForegroundColor Green
$tenantID= Read-Host -Prompt "Tenant ID"


If(![string]::IsNullOrEmpty($tenantID)) {
    Connect-MicrosoftTeams -TenantId $tenantID
    if($useOfMG -eq "y") {
        Connect-MgGraph -Scopes "User.Read.All","Group.Read.All","TeamSettings.Read.All" -TenantId $tenantID -NoWelcome
    }
    
} else {
    Connect-MicrosoftTeams
    if($useOfMG -eq "y") {
        Connect-MgGraph -Scopes "User.Read.All","Group.Read.All","TeamSettings.Read.All" -NoWelcome
    }
}

function ReadCSVFile {
    try {
        $Global:csv = import-csv $path -Delimiter ";"
    }
    catch {
        write-host "CSV could not be imported" -ForegroundColor Red
        $error[0].Exception
        break
    }
}

$fileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'Comma Seperated File (*.csv)|*.csv'
}
Write-Host "Please choose the a users csv file."
$null = $fileBrowser.ShowDialog()

if([string]::IsNullOrEmpty($fileBrowser.FileName)) {
    break
} else {
    $path = $fileBrowser.FileName
    ReadCSVFile
}

$sortedCSV = $csv | Where-Object { $_.Generated -ne 'y'}

$sortedCSV | ForEach-Object {
    $name = $_.Name
    $resourceAccount = $_.'Resource Account'
    $phoneNumber = $_.'Phone Number'
    $teamsName = $_.'Teams Name'
    $channelName = $_.'Channel Name'
    $channelUserObject = $_.'Teams Owner'
    $language = $_.'Language'
    $dialPlan = $_.'Dial Plan'
    $voiceRoutingPolicy = $_.'Voice Routing Policy'
    $routingMethod = $_.'Routing Method'
    $presenceBasedRouting = (GetBool -onoff $_.'Presence-based routing')
    $conferenceMode = (GetBool -onoff $_.'Conference Mode')
    $canOptOut = (GetBool -onoff $_.'Can Opt-Out')
    $agentAlertTime = $_.'Agent Alert time (seconds)'
    $callOverflowMaxCall = $_.'Call overflow max call'
    $callOverflowAction = $_.'Call overflow action'
    $callOverflowTarget = $_.'Call overflow target'
    $callTimeout = $_.'Call timeout'
    $callTimeoutAction = $_.'Call timeout action'
    $callTimeoutTarget = $_.'Call timeout target'
    $noAgentsOnlyOnNewCalls = (GetBool -onoff $_.'No agents only on new calls')
    $noAgentsSelection = $_.'No agents action'
    $noAgentsTarget = $_.'No agents target'
    $callAnswering = $_.'Call answering'

    Write-Host "Creating call queue $name ..." -NoNewline

    $callOverflowTargetID = $null
    $callTimeoutTargetID = $null
    $noAgentsTargetID = $null

    $callQueueID = Get-CSCallQueue -NameFilter "$name" | Select-Object Identity
    if(![string]::IsNullOrEmpty($callQueueID)) {
        Write-Host "Call Queue $name already exists, skipping!" -ForegroundColor DarkYellow
        Return
    }

    # Call Overflow Target
    if(![string]::IsNullOrEmpty($callOverflowTarget)) {
        if($useOfMG -eq "y") {
            $callOverflowTargetID = Get-MGUser -Filter "UserPrincipalName eq '$callOverflowTarget'" | Select-Object -ExpandProperty ID
        } else {
            $callOverflowTargetID = Get-CsOnlineUser -Filter {UserPrincipalName -eq "$callOverflowTarget"} | Select-Object -ExpandProperty Identity
        }
    }

    # Call Timeout Target
    if(![string]::IsNullOrEmpty($callTimeoutTarget)) {
        if($useOfMG -eq "y") {
            $callTimeoutTargetID = Get-MGUser -Filter "UserPrincipalName eq '$callTimeoutTarget'" | Select-Object -ExpandProperty ID
        } else {
            $callTimeoutTargetID = Get-CsOnlineUser -Filter {UserPrincipalName -eq "$callTimeoutTarget"} | Select-Object -ExpandProperty Identity
        }
    }

    # No Agents Target
    if(![string]::IsNullOrEmpty($noAgentsTarget)) {
        if($useOfMG -eq "y") {
            $noAgentsTargetID = Get-MGUser -Filter "UserPrincipalName eq '$noAgentsTarget'" | Select-Object -ExpandProperty ID
        } else {
            $noAgentsTargetID = Get-CsOnlineUser -Filter {UserPrincipalName -eq "$noAgentsTarget"} | Select-Object -ExpandProperty Identity
        }
    }

    $parameters = @{
        Name = $name
        RoutingMethod = $routingMethod
        UseDefaultMusicOnHold = $true
        LanguageId = $language
        AllowOptOut = $canOptOut
        AgentAlertTime = $agentalerttime
        ConferenceMode = $conferenceMode
        OverflowThreshold = $callOverflowMaxCall
        OverflowAction = $callOverflowAction
        OverflowActionTarget = $callOverflowTargetID
        PresenceBasedRouting = $presenceBasedRouting
        TimeoutThreshold = $callTimeout
        TimeoutAction = $callTimeoutAction
        TimeoutActionTarget = $callTimeoutTargetID
        NoAgentAction = $noAgentsSelection
        NoAgentActionTarget = $noAgentsTargetID
        ShouldOverwriteCallableChannelProperty = $true
    }

    # Call Anwswering
    if($callAnswering -eq "Team") {
        # Get Channel Owner ID
        if(![string]::IsNullOrEmpty($channelUserObject)) {
            if($useOfMG -eq "y") {
                $channelUserObjectID = Get-MGUser -Filter "UserPrincipalName eq '$channelUserObject'" | Select-Object -ExpandProperty ID
            } else {
                $channelUserObjectID = Get-CsOnlineUser -Filter {UserPrincipalName -eq "$channelUserObject"} | Select-Object -ExpandProperty Identity
            }
        }
        if($useOfMG -eq "y") {
            $teamID = Get-MgTeam -Filter "DisplayName eq '$teamName'" | Select-Object -ExpandProperty ID
        } else {
            $teamID = Get-Team -DisplayName "$teamName" | Select-Object -ExpandProperty GroupID
        }
        $channelID = Get-TeamAllChannel -GroupId $teamID | Where {$_.DisplayName -eq $channelName} | Select -ExpandProperty ID
        $parameters.Add('ChannelId',$channelID)
        $parameters.Add('ChannelUserObjectId',$channelUserObjectID)
        $parameters.Add('DistributionLists',$teamID)
    } elseif ($callAnswering -eq "Users") {
        <# Action when this condition is true #>
    } elseif ($callAnswering -eq "Groups") {
        <# Action when this condition is true #>
    }

    try {
        $newCallQueueID = New-CSCallQueue @parameters
    }
    catch {
        Write-host -f Red "error creating Call Queue :" $_.Exception.Message
    }

    
    if($useOfMG -eq "y") {
        $resourceAccountID = Get-MGUser -Filter "UserPrincipalName eq '$resourceAccount'" | Select-Object -ExpandProperty ID
    } else {
        $resourceAccountID = Get-CsOnlineUser -Filter {UserPrincipalName -eq "$resourceAccount"} | Select-Object -ExpandProperty Identity
    }

    if(![string]::IsNullOrEmpty($resourceAccountID) -and ![string]::IsNullOrEmpty($newCallQueueID)) {
        try {
            New-CsOnlineApplicationInstanceAssociation -Identities @($resourceAccountID) -ConfigurationID $newCallQueueID -ConfigurationType CallQueue
        }
        catch {
            Write-host -f Red "error creating Call Queue :" $_.Exception.Message
        }
        if(![string]::IsNullOrEmpty($phoneNumber)) {
            try {
                Set-CsPhoneNumberAssignment -Identity $resourceAccount -PhoneNumber $phoneNumber -PhoneNumberType DirectRouting
            }
            catch {
                Write-host -f Red "error creating Call Queue :" $_.Exception.Message
            }
        } else {
            Write-host "success!" -f Green
        }
    } else {
        Write-host "success!" -f Green
    }

}

Write-Host "All Call Queues created successfully." -ForegroundColor Green
Pause