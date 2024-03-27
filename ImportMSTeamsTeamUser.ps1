# ImportMSTeamsTeamUsers.ps1
# Use CSV file to users to Teams.
#
# REQUIREMENTS:
# PS-Module Microsoft-Teams and Microsoft.Graph (Optional)
# Graph-Api authorizations needed: User.Read.All, Group.Read.All, TeamSettings.Read.All
# Export Worksheet Teams from MSTeams-Runbook.xlsx as CSV file
#
# Version 1.0.2 (Build 1.0.1-2024-03-27)
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

$sortedCSV = $csv | Where-Object { $_.Generated -ne 'y'}

$sortedCSV | ForEach-Object {
    $teamName = $_.'Teams Name'
    $upn = $_.User
    $role = $_.Role

    Write-Host "Adding user $upn to $teamName ..." -NoNewline

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

    if(![string]::IsNullOrEmpty($teamID)) {
        try {
            Add-TeamUser -GroupId $teamID -User $upn -Role $role
            Write-host "success!" -f Green
        }
        catch {
            Write-host -f Red "error Adding $upn :" $_.Exception.Message
        }
    } else {
        Write-Host "$teamName doesn't exists, skipping!" -ForegroundColor DarkYellow
    }
}

Write-Host "All users added to their teams was successful." -ForegroundColor Green
Pause
Write-Host "Enter the UPN of the user who run this script to remove as owner" -ForegroundColor Blue
$adminUPN = Read-Host "Admin UPN"

$sortedCSV = $csv | Where-Object { $_.Generated -ne 'y'} | Select-Object -Unique -Property 'Team'

$sortedCSV | ForEach-Object {
    $teamName = $_.Team

    Write-Host "Removing owner $adminUPN from $teamName ..." -NoNewline

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

    if(![string]::IsNullOrEmpty($teamID)) {
        try {
            Remove-TeamUser -GroupId $teamID -User $adminUPN
            Write-host "success!" -f Green
        }
        catch {
            Write-host -f Red "error removing $adminUPN :" $_.Exception.Message
        }
    } else {
        Write-Host "$teamName doesn't exists, skipping!" -ForegroundColor DarkYellow
    }
}