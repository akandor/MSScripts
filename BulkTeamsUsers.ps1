# BulkTeamsUsers.ps1
# Use CSV file of users to create users in bulk.
#
# REQUIREMENTS:
# License needs to be assigned before
# Export Worksheet Users from MSTeams-Runbook.xlsx as CSV file
#
# Version 1.0.0 (Build 1.0.0-2023-12-14)
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
} else {
    Connect-MicrosoftTeams
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

#CSV Columns
#UPN                  
#Number               
#Calling Policy       
#Dial Plan            
#Voice Routing Policy 
#Call Hold Policy      
#Call Park Policy      
#Caller ID Policy      
#Emergency Policy     
#Voice Mail           
#Call Queue           
#Generated            

$i = 0
$users = @()
$sortedCSV = $csv | Where-Object { $_.Generated -ne 'y'}

foreach ($row in $sortedCSV) {
    $users += $row
}

foreach ($user in $users) {

    $progress = 100 / ($users.Count) * ($i)
    $progressRounded = [math]::Round($progress)

    $upn = $user.UPN
    $number = $user.Number
    $callingPolicy = $user.'Calling Policy'
    $dialPlan = $user.'Dial Plan'
    $voiceRoutingPolicy = $user.'Voice Routing Policy'
    $callHoldPolicy = $user.'Call Hold Policy'
    $callParkPolicy = $user.'Call Park Policy'
    $callerIdPolicy = $user.'Caller ID Policy'
    $emergencyCallRoutingPolicy = $user.'Emergency Policy'
    $voiceMailLanguage = $user.'Voice Mail'

    Write-Progress -Activity "Creating User $upn" -Id 1 -Status "$progressRounded% Complete" -PercentComplete $progress

    $i++

    if(![string]::IsNullOrEmpty($number)) {
        try {
            Set-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $number -PhoneNumberType DirectRouting -ErrorAction Stop
        }
        catch {
            write-host "Cannot assign number $number to $upn" -ForegroundColor Red
            $error[0].Exception
            continue
        }
    } else {
        try {
            Set-CsPhoneNumberAssignment -Identity $upn -EnterpriseVoiceEnabled $true -ErrorAction Stop
        }
        catch {
            write-host "Cannot enable Enterprise Voice on $upn" -ForegroundColor Red
            $error[0].Exception
            continue
        }
    }

    if(![string]::IsNullOrEmpty($callingPolicy)) {
        try {
            Grant-CsTeamsCallingPolicy -Identity $upn -PolicyName $callingPolicy -ErrorAction Stop
        } catch {
            write-host "Cannot assign $callingPolicy to $upn" -ForegroundColor Red
            $error[0].Exception
            continue
        }
    }
    if(![string]::IsNullOrEmpty($dialPlan)) {
        try {
            Grant-CsTenantDialPlan -Identity $upn -PolicyName $dialPlan -ErrorAction Stop
        } catch {
            write-host "Cannot assign $dialPlan to $upn" -ForegroundColor Red
            $error[0].Exception
            continue
        }
    }
    if(![string]::IsNullOrEmpty($voiceRoutingPolicy)) {
        try {
            Grant-CsOnlineVoiceRoutingPolicy -Identity $upn -PolicyName $voiceRoutingPolicy -ErrorAction Stop
        } catch {
            write-host "Cannot assign $voiceRoutingPolicy to $upn" -ForegroundColor Red
            $error[0].Exception
            continue
        }
    }
    if(![string]::IsNullOrEmpty($callHoldPolicy)) {
        try {
            Grant-CsTeamsCallHoldPolicy -Identity $upn -PolicyName $callHoldPolicy -ErrorAction Stop
        } catch {
            write-host "Cannot assign $callHoldPolicy to $upn" -ForegroundColor Red
            $error[0].Exception
            continue
        }
    }
    if(![string]::IsNullOrEmpty($callParkPolicy)) {
        try {
            Grant-CsTeamsCallParkPolicy -Identity $upn -PolicyName $callParkPolicy -ErrorAction Stop
        } catch {
            write-host "Cannot assign $callParkPolicy to $upn" -ForegroundColor Red
            $error[0].Exception
            continue
        }
    }
    if(![string]::IsNullOrEmpty($callerIdPolicy)) {
        try {
            Grant-CsCallingLineIdentity -Identity $upn -PolicyName $callerIdPolicy -ErrorAction Stop
        } catch {
            write-host "Cannot assign $callerIdPolicy to $upn" -ForegroundColor Red
            $error[0].Exception
            continue
        }
    }
    if(![string]::IsNullOrEmpty($emergencyCallRoutingPolicy)) {
        try {
            Grant-CsTeamsEmergencyCallRoutingPolicy -Identity $upn -PolicyName $emergencyCallRoutingPolicy -ErrorAction Stop
        } catch {
            write-host "Cannot assign $emergencyCallRoutingPolicy to $upn" -ForegroundColor Red
            $error[0].Exception
            continue
        }
    }
    if(![string]::IsNullOrEmpty($voiceMailLanguage)) {
        try {
            Set-CsOnlineVoicemailUserSettings -Identity $upn -PromptLanguage $voiceMailLanguage -ErrorAction Stop
        } catch {
            write-host "Cannot assign $voiceMailLanguage to $upn" -ForegroundColor Red
            $error[0].Exception
            continue
        }
    }

}

Write-Progress -Activity "Creating Users completed!" -Id 1 -Completed
Pause
Exit