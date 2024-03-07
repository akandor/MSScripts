# NewTeamsUser.ps1
# Reads out the policies and you can create a new MS Teams user.
# REQUIREMENTS: License needs to be assigned before
#
# Version 1.0.1 (Build 1.0.1-2024-01-21)
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

do {
    Clear-Host
    Write-Host "==================================="
    Write-Host ""
    Write-Host "     Create New MS Teams User"
    Write-Host ""
    Write-Host "==================================="
    Write-Host ""
    do {
        $upn = Read-Host "Please enter the UserPrincipalName"
    } while ([string]::IsNullOrEmpty($upn))

    do {
        $number = Read-Host "Please enter the Phone number (E.164 Format)"
    } while ([string]::IsNullOrEmpty($number))
    Clear-Host
    Write-Host ""
    Write-Host "    Voice Routing Policies"
    Write-Host ""
    $voiceRoutingPolicies = Get-CsOnlineVoiceRoutingPolicy

    $i = 1
    foreach($voiceRoutingPolicy in $voiceRoutingPolicies) {
        $vrp_identity = $voiceRoutingPolicy.Identity -replace "Tag:",""
        Write-Host "[$i] $vrp_identity"
        $i++
    }
    do {
        $vrp_no = Read-Host "Please choose a Voice Routing Policy"
    } while ([string]::IsNullOrEmpty($vrp_no))

    $used_vrp = $voiceRoutingPolicies[$vrp_no-1].Identity -replace "Tag:",""

    Clear-Host
    Write-Host ""
    Write-Host "    Calling Policies"
    Write-Host ""
    $callingPolicies = Get-CsTeamsCallingPolicy

    $i = 1
    foreach ($callingPolicy in $callingPolicies) {
        $cp_identity = $callingPolicy.Identity -replace "Tag:",""
        Write-Host "[$i] $cp_identity"
        $i++
    }
    do {
        $cp_no = Read-Host "Please choose a Calling Policy"
    } while ([string]::IsNullOrEmpty($cp_no))

    $used_cp = $callingPolicies[$cp_no-1].Identity -replace "Tag:",""

    Clear-Host
    Write-Host ""
    Write-Host "    Dial Plan"
    Write-Host ""
    $dialPlans = Get-CsTenantDialPlan

    $i = 1
    foreach ($dialPlan in $dialPlans) {
        $dp_identity = $dialPlan.Identity -replace "Tag:",""
        Write-Host "[$i] $dp_identity"
        $i++
    }
    do {
        $dp_no = Read-Host "Please choose a Dial Plan"
    } while ([string]::IsNullOrEmpty($dp_no))

    $used_dp = $dialPlans[$dp_no-1].Identity -replace "Tag:",""

    Clear-Host
    Write-Host ""
    Write-Host "    Call Hold Policy"
    Write-Host ""
    $callHoldPolicies = Get-CsTeamsCallHoldPolicy

    $i = 1
    foreach ($callHoldPolicy in $callHoldPolicies) {
        $chp_identity = $callHoldPolicy.Identity -replace "Tag:",""
        Write-Host "[$i] $chp_identity"
        $i++
    }
    do {
        $chp_no = Read-Host "Please choose a Call Hold Policy"
    } while ([string]::IsNullOrEmpty($chp_no))

    $used_chp = $callHoldPolicies[$chp_no-1].Identity -replace "Tag:",""

    Clear-Host
    Write-Host ""
    Write-Host "    Emergency Call Routing Policy"
    Write-Host ""
    $emergencyCallRoutingPolicies = Get-CsTeamsEmergencyCallRoutingPolicy

    $i = 1
    foreach ($emergencyCallRoutingPolicy in $emergencyCallRoutingPolicies) {
        $ecrp_identity = $emergencyCallRoutingPolicy.Identity -replace "Tag:",""
        Write-Host "[$i] $ecrp_identity"
        $i++
    }
    do {
        $ecrp_no = Read-Host "Please choose a Emergency Call Routing Policy"
    } while ([string]::IsNullOrEmpty($chp_no))

    $used_ecrp = $emergencyCallRoutingPolicies[$ecrp_no-1].Identity -replace "Tag:",""

    Clear-Host
    Write-Host "==================================="
    Write-Host ""
    Write-Host "          Check Values"
    Write-Host ""
    Write-Host "==================================="
    Write-Host ""
    Write-Host "UserPrincipalName: $upn"
    Write-Host "Phone Number: $number"
    Write-Host "Voice Routing Policy: $used_vrp"
    Write-Host "Calling Policy: $used_cp"
    Write-Host "Dial Plan: $used_dp"
    Write-Host "Call Hold Policiy: $used_chp"
    Write-Host "Emergency Call Routing Policy: $used_ecrp"
    Write-Host "----------------------------------------------"
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', 'All values are correct. Proceed to create user.'
    $no = New-Object System.Management.Automation.Host.ChoiceDescription '&No', 'Value not correct. Start over!'
    $quit = New-Object System.Management.Automation.Host.ChoiceDescription '&Quit', 'Quit the script without creating the user.'
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no, $quit)

    $message = 'Are the values correct?'
    $result = $host.ui.PromptForChoice($null, $message, $options, 0)
}
while($result -eq "1")

if($result -eq "2") {
    Exit
}

if($result -eq "0") {
    # Create User
    try {
        Set-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $number -PhoneNumberType DirectRouting -ErrorAction Stop
        Grant-CsTenantDialPlan -Identity $upn -PolicyName $used_dp -ErrorAction Stop
        Grant-CsOnlineVoiceRoutingPolicy -Identity $upn -PolicyName $used_vrp -ErrorAction Stop
        Grant-CsTeamsCallingPolicy -Identity $upn -PolicyName $used_cp -ErrorAction Stop #No resource
        Grant-CsTeamsCallHoldPolicy -Identity $upn -PolicyName $used_chp -ErrorAction Stop #No Resource
        Grant-CsTeamsEmergencyCallRoutingPolicy -Identity $upn -PolicyName $used_ecrp -ErrorAction Stop #No resource
        Write-Host "User $upn created successfully!" -ForegroundColor Green
    }
    catch {
        write-host "Cannot create user $upn" -ForegroundColor Red
        write-host "Reason:"
        $error[0].Exception
        continue
    }
}

Pause 