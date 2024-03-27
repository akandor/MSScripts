# CreateMSTeamsResourceAccounts.ps1
# Use CSV file to create resource accounts and assign license.
#
# REQUIREMENTS:
# PS-Module Microsoft-Teams and Microsoft.Graph
# Graph-Api authorizations needed: User.ReadWrite.All, Group.Read.All, TeamSettings.Read.All, Organization.Read.All
# Export Worksheet Call Queue from MSTeams-Runbook.xlsx as CSV file
#
# Version 1.0.0 (Build 1.0.0-2024-03-26)
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
Write-Host "Please enter your Tenant ID (can be left empty if you have only one)" -ForegroundColor Green
$tenantID= Read-Host -Prompt "Tenant ID"

If(![string]::IsNullOrEmpty($tenantID)) {
    Connect-MicrosoftTeams -TenantId $tenantID
    Connect-MgGraph -Scopes "User.ReadWrite.All","Group.Read.All","TeamSettings.Read.All","Organization.Read.All" -TenantId $tenantID
} else {
    Connect-MicrosoftTeams
    Connect-MgGraph -Scopes "User.ReadWrite.All","Group.Read.All","TeamSettings.Read.All","Organization.Read.All"
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

# Get license information
# for every Call Queue a phone system license is needed
$phoneVirtualSku = Get-MgSubscribedSku -All | Where SkuPartNumber -eq 'PHONESYSTEM_VIRTUALUSER'

$sortedCSV = $csv | Where-Object { $_.Generated -ne 'y'}

$sortedCSV | ForEach-Object {
    $name = $_.Name
    $resourceAccount = $_.'Resource Account'
    $usageLocation = $_.'Usage Location'

    Write-Host "Creating resource account $name ..." -NoNewline

    try {
        New-CsOnlineApplicationInstance -UserPrincipalName $resourceAccount -DisplayName $name -ApplicationID "11cd3e2e-fccb-42ad-ad00-878b93575e07"
        Write-host "success!" -f Green
    }
    catch {
        Write-host -f Red "error creating resource account :" $_.Exception.Message
    }

    Write-Host "Wait for 5 seconds to sync..." -ForegroundColor Magenta
    Start-Sleep -Seconds 5

    Write-Host "Adding license to resource account $name ..." -NoNewline

    try {
        Update-MgUser -UserId $resourceAccount -UsageLocation $usageLocation
        Set-MgUserLicense -UserId "$resourceAccount" -AddLicenses @{SkuId = $phoneVirtualSku.SkuId} -RemoveLicenses @()
        Write-host "success!" -f Green
    }
    catch {
        Write-host -f Red "error adding license to resource account :" $_.Exception.Message
    }
}