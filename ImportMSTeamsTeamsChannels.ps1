# ImportMSTeamsTeamsChannels.ps1
# Use CSV file to create Teams and Channels.
#
# REQUIREMENTS:
# PS-Module Microsoft-Teams and Microsoft.Graph (Optional)
# Graph-Api authorizations needed: User.Read.All, Group.Read.All, TeamSettings.Read.All
# Export Worksheet Call Queues from MSTeams-Runbook.xlsx as CSV file
#
# Version 1.0.2 (Build 1.0.2-2024-03-27)
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

$sortedCSV = $csv | Where-Object { $_.Generated -ne 'y'} | Select-Object -Unique -Property 'Teams Name'

$sortedCSV | ForEach-Object {
    $teamName = $_.'Teams Name'

    Write-Host "Creating Team: $teamName ..." -NoNewline
    #Check if Team exists
    try {
        if($useOfMG -eq "y") {
            $teamID = Get-MgTeam -Filter "DisplayName eq '$teamName'" | Select-Object -ExpandProperty ID
        } else {
            $teamID = Get-Team -DisplayName "$teamName" | Select-Object -ExpandProperty GroupID
        }
    }
    catch {
        Write-host -f Red "error getting GroupID of $teamName :" $_.Exception.Message
    }
    
    if([string]::IsNullOrEmpty($teamID)) {
        #Create Team
        try {
            $teamID = New-Team -DisplayName $teamName -Visibility Public -AllowGuestCreateUpdateChannels $false -AllowGuestDeleteChannels $false -AllowCreateUpdateChannels $false -AllowDeleteChannels $false -AllowAddRemoveApps $false -AllowCreateUpdateRemoveTabs $false -AllowCreateUpdateRemoveConnectors $false 
            Write-Host "ID "$teamID.Id "success." -ForegroundColor Green
        }
        catch {
            Write-host -f Red "error creating $teamName :" $_.Exception.Message
        }
        
    } else {
        Write-Host "ID $teamID already exists, skipping!" -ForegroundColor DarkYellow
    }
}

Write-Host "All Teams created successfully. In the next step channels will be created" -ForegroundColor Green
Pause

$sortedCSV = @()
$sortedCSV = $csv | Where-Object { $_.Generated -ne 'y'}

$sortedCSV | ForEach-Object {
    $teamName = $_.'Teams Name'
    $channelName = $_.'Channel Name'

    Write-Host "Creating Channel: $channelName ..." -NoNewline

    try {
        if($useOfMG -eq "y") {
            $teamID = Get-MgTeam -Filter "DisplayName eq '$teamName'" | Select-Object -ExpandProperty ID
        } else {
            $teamID = Get-Team -DisplayName "$teamName" | Select-Object -ExpandProperty GroupID
        }
    }
    catch {
        Write-host -f Red "error getting ID of $teamName :" $_.Exception.Message
    }

    try {
        $channelID = Get-TeamAllChannel -GroupId $teamID | Where {$_.DisplayName -eq $channelName} | Select -ExpandProperty ID
    }
    catch {
        Write-host -f Red "error getting ID of $channelName :" $_.Exception.Message
    }
    
    if([string]::IsNullOrEmpty($channelID)) {
        try {
            $channelID = New-TeamChannel -GroupId $teamID -DisplayName $channelName
            Write-Host "ID " $channelID.Id "success." -ForegroundColor Green
        }
        catch {
            Write-host -f Red "error creating $channelName :" $_.Exception.Message
        }
    } else {
        Write-Host "ID " $channelID "already exists, skipping!" -ForegroundColor DarkYellow
    }
}

Write-Host "All channels created successfully." -ForegroundColor Green
Pause