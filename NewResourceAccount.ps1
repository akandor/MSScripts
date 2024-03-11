# NewResourceAccount.ps1
# Reads out the policies and you can add a new number to a resource account.
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

function Load-Module ($m) {

    # If module is imported say that and do nothing
    if (Get-Module | Where-Object {$_.Name -eq $m}) {
        write-host "Module $m is already imported."
    }
    else {

        # If module is not imported, but available on disk then import
        if (Get-Module -ListAvailable | Where-Object {$_.Name -eq $m}) {
            Import-Module $m -Verbose
        }
        else {

            # If module is not imported, not available on disk, but is in online gallery then install and import
            if (Find-Module -Name $m | Where-Object {$_.Name -eq $m}) {
                Install-Module -Name $m -Force -Verbose -Scope CurrentUser
                Import-Module $m -Verbose
            }
            else {

                # If the module is not imported, not available and not in the online gallery then abort
                write-host "Module $m not imported, not available and not in an online gallery, exiting."
                EXIT 1
            }
        }
    }
}

Load-Module "MicrosoftTeams"

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
    Write-Host "   New MS Teams Resource Account"
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
    Write-Host "==================================="
    Write-Host ""
    Write-Host "          Check Values"
    Write-Host ""
    Write-Host "==================================="
    Write-Host ""
    Write-Host "UserPrincipalName: $upn"
    Write-Host "Phone Number: $number"
    Write-Host "Voice Routing Policy: $used_vrp"
    Write-Host "Dial Plan: $used_dp"
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
        Write-Host "Resource Account $upn updated successfully!" -ForegroundColor Green
    }
    catch {
        write-host "Cannot update resource account $upn" -ForegroundColor Red
        write-host "Reason:"
        $error[0].Exception
        continue
    }
}

Pause 