# OffboardTeamsUser.ps1
# Removes number from Teams User and optinally remove EnterpriseVoice.
#
# Version 1.0.0 (Build 1.0.0-2024-02-28)
# 
# Created by: Armin Toepper
#
#########################################################################################
#
#########################################################################################
#                            DO NOT EDIT BELOW THESE LINES!                             #
#########################################################################################
#
Set-StrictMode -Version "2.0"
$ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent 

Clear-Host
Write-Host "Please enter your Tenant ID (can be left empty if you have only one)" -ForegroundColor Green
$tenantID= Read-Host -Prompt "Tenant ID"
If(![string]::IsNullOrEmpty($tenantID)) {
    Connect-MicrosoftTeams -TenantId $tenantID
} else {
    Connect-MicrosoftTeams
}

Clear-Host
Write-Host "==================================="
Write-Host ""
Write-Host "     Offboard MS Teams User"
Write-Host ""
Write-Host "==================================="
Write-Host ""
do {
    $upn = Read-Host "Please enter the UserPrincipalName"
} while ([string]::IsNullOrEmpty($upn))

$yes = New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', 'Choose Yes to set the user enterprise voice enables to false.'
$no = New-Object System.Management.Automation.Host.ChoiceDescription '&No', 'Choose no to remove just the number'
$quit = New-Object System.Management.Automation.Host.ChoiceDescription '&Quit', 'Quit the script without offboarding the user.'
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no, $quit)

$message = 'Do you want to remove EnterpriseVoiceEnabled too?'
$result = $host.ui.PromptForChoice($null, $message, $options, 0)

if($result -eq "0") {
    try {
        Remove-CsPhoneNumberAssignment -Identity $upn -RemoveAll -ErrorAction Stop
        Set-CsPhoneNumberAssignment -Identity $upn -EnterpriseVoiceEnabled $false -ErrorAction Stop
        Write-Host "Offboard User $upn and disabled EnterpriseVoice successful." -ForegroundColor Green
    }
    catch {
        write-host "Cannot offboard user $upn" -ForegroundColor Red
        write-host "Reason:"
        $error[0].Exception
        continue
    }
    Pause
} elseif($result -eq "1") {
    try {
        Remove-CsPhoneNumberAssignment -Identity $upn -RemoveAll -ErrorAction Stop
        Write-Host "Offboard User $upn successful." -ForegroundColor Green
    }
    catch {
        write-host "Cannot offboard user $upn" -ForegroundColor Red
        write-host "Reason:"
        $error[0].Exception
        continue
    }
    Pause
} else {
    Exit
}