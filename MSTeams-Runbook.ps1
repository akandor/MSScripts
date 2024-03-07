####
# Filename:
#   MSTeams-Runbook.ps1
#
# Author:
#   Armin Toepper
#
# Last changes:
#   2024-02-14 first release
#   latest updates found in ReleaseNotes.txt
#
#
# Description:
#   Create or update MS Teams policies and dialplan based on a runbook.
#
#
####
Set-StrictMode -Version "2.0"

$ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

Add-Type -AssemblyName System.Windows.Forms
# Import Local Copy of ImportExcel Module
try {
    Import-Module -Name ($ScriptDirectory + "/library/ImportExcel/ImportExcel")
}
catch {
    Write-Host "Error while loading supporting PowerShell Scripts"
    break
}

Function LogWrite
{
   Param (
    [string]$logstring
    )
   $DateTime = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
   $LogMessage = "$Datetime $LogString"
   Add-content $Logfile -value $LogMessage
}

function ReadExcelFile {
    try {
        $excelFile = Import-Excel $excelpath -WorksheetName $workSheetName
    }
    catch {
        Write-Host "Error while loading Excel File"
        break
    }
    return $excelFile
}

function CreateSBC {
    foreach ($row in $excelData) {
        $parameters = @{
            Fqdn = $row.Fqdn
            Enabled = (GetBool -onoff $row.Enabled)
            SipSignalingPort = $row.SipSignalingPort
            SendSipOptions = (GetBool -onoff $row.SendSipOptions)
            ForwardCallHistory = (GetBool -onoff $row.ForwardCallHistory)
            ForwardPai = (GetBool -onoff $row.ForwardPai)
            MaxConcurrentSessions = $row.MaxConcurrentSessions
            FailoverResponseCodes = $row.FailoverResponseCodes
            FailoverTimeSeconds = $row.FailoverTimeSeconds
            PidfLoSupported = (GetBool -onoff $row.PidfloSupported)
        }
        $updateParameters = @{
            Identity = $row.Fqdn
            Enabled = (GetBool -onoff $row.Enabled)
            SipSignalingPort = $row.SipSignalingPort
            SendSipOptions = (GetBool -onoff $row.SendSipOptions)
            ForwardCallHistory = (GetBool -onoff $row.ForwardCallHistory)
            ForwardPai = (GetBool -onoff $row.ForwardPai)
            MaxConcurrentSessions = $row.MaxConcurrentSessions
            FailoverResponseCodes = $row.FailoverResponseCodes
            FailoverTimeSeconds = $row.FailoverTimeSeconds
            PidfLoSupported = (GetBool -onoff $row.PidfloSupported)
        }

        $getSBC = Get-CsOnlinePSTNGateway -Identity $row.Fqdn -ErrorAction SilentlyContinue

        if(![string]::IsNullOrEmpty($getSBC)) {
            Clear-Host
            Write-Host "SBC already exists!" -ForegroundColor Yellow
            Write-Host $row.Fqdn -ForegroundColor Yellow
            $choose = Read-Host "Do you want to update the settings? [y/N]"
            if ($choose -eq "y") {
                Set-CsOnlinePSTNGateway @updateParameters
                Write-Host ""
                Write-Host "SBC updated successfully!" -ForegroundColor Green
                Write-Host ""
                pause
            }
        } else {
            New-CsOnlinePSTNGateway @parameters
        }  
    }
    Write-Host ""
    Write-Host "SBCs done!" -ForegroundColor Green
    Write-Host ""
    pause
}

function CreatePSTNUsages {
    Clear-Host
    foreach ($row in $excelData) {
        $pstnUsage = $row.Name

        try {
            Set-CsOnlinePstnUsage -Identity global -Usage @{add=$pstnUsage} -ErrorAction  SilentlyContinue
        }
        catch {
        }
    }
    Write-Host ""
    Write-Host "PSTN usages created."
    Write-Host ""
    pause
}

function CreateVoiceRoutes {
    foreach ($row in $excelData) {
        $parameters = @{
            Identity = $row.Name
            Description = $row.Description
            NumberPattern = $row.NumberPattern
            OnlinePstnGatewayList = $row.OnlinePSTNGatewayList
            Priority = $row.Priority
            OnlinePstnUsages = $row.OnlinePstnUsages
        }

        $getRoute = Get-CsOnlineVoiceRoute -Identity $row.Name -ErrorAction SilentlyContinue

        if(![string]::IsNullOrEmpty($getRoute)) {
            Clear-Host
            Write-Host "Voice Route already exists!" -ForegroundColor Yellow
            Write-Host $row.Name -ForegroundColor Yellow
            $choose = Read-Host "Do you want to update the settings? [y/N]"
            if ($choose -eq "y") {
                Set-CsOnlineVoiceRoute @parameters
                Write-Host ""
                Write-Host "Voice route updated successfully!" -ForegroundColor Green
                Write-Host ""
                pause
            }
        } else {
            New-CsOnlineVoiceRoute @parameters
        } 
    }
    Write-Host ""
    Write-Host "Voice routes done!" -ForegroundColor Green
    Write-Host ""
    pause
}

function CreateVoiceRoutingPolicies {
    foreach ($row in $excelData) {

        $getVoiceRoute = Get-CsOnlineVoiceRoutingPolicy -Identity $row.Name -ErrorAction SilentlyContinue

        if(![string]::IsNullOrEmpty($getVoiceRoute)) {
            Clear-Host
            Write-Host "Voice routing policy already exists!" -ForegroundColor Yellow
            Write-Host $row.Name -ForegroundColor Yellow
            $choose = Read-Host "Do you want to update the settings? [y/N]"
            if ($choose -eq "y") {
                $listPstnUsage = @()
                $listPstnKey = $row | Get-Member -Name "PSTN usage records*"
                foreach ($pstnKey in $listPstnKey) {
                    $pstnValue = $row.($pstnKey.Name)
                    if (![string]::IsNullOrEmpty($pstnValue)) {
                        $listPstnUsage += $pstnValue
                    }
                }
                Set-CsOnlineVoiceRoutingPolicy -Identity $row.Name -Description $row.Description -OnlinePstnUsages $listPstnUsage
                Write-Host ""
                Write-Host "Voice routing policy updated successfully!" -ForegroundColor Green
                Write-Host ""
                pause
            }
        } else {
            $listPstnUsage = @()
            $listPstnKey = $row | Get-Member -Name "PSTN usage records*"
            foreach ($pstnKey in $listPstnKey) {
                $pstnValue = $row.($pstnKey.Name)
                if (![string]::IsNullOrEmpty($pstnValue)) {
                    $listPstnUsage += $pstnValue
                }
            }
            New-CsOnlineVoiceRoutingPolicy -Identity $row.Name -Description $row.Description -OnlinePstnUsages $listPstnUsage
        }     
    }
    Write-Host ""
    Write-Host "Voice routing policies done!" -ForegroundColor Green
    Write-Host ""
    pause
}

function CreateEmergencyCallRoutingPolicies {
    $emergencyData = $excelData | Group-Object -Property Name

    foreach ($data in $emergencyData) {
        
        $getECRP = Get-CsTeamsEmergencyCallRoutingPolicy -Identity $data.Name -ErrorAction SilentlyContinue

        if(![string]::IsNullOrEmpty($getECRP)) {
            Clear-Host
            Write-Host "Emergency call routing policy already exists!" -ForegroundColor Yellow
            Write-Host $data.Name -ForegroundColor Yellow
            $choose = Read-Host "Do you want to update the settings? [y/N]"
            if ($choose -eq "y") {
                $emergencyNumbers = @()

                foreach ($row in $data.Group) {
                    $emergencyNumbers += New-CsTeamsEmergencyNumber -EmergencyDialString $row.EmergencyDialString -EmergencyDialMask $row.EmergencyDialMask -OnlinePSTNUsage $row.OnlinePSTNUsage
                }
                Set-CsTeamsEmergencyCallRoutingPolicy -Identity $row.Name -Description $row.Description -AllowEnhancedEmergencyServices (GetBool -onoff $row.DynamicEmergencyCalling) -EmergencyNumbers $emergencyNumbers
                Write-Host ""
                Write-Host "Emergency call routing policy updated successfully!" -ForegroundColor Green
                Write-Host ""
                pause
            }
        } else {
            $emergencyNumbers = @()

            foreach ($row in $data.Group) {
                $emergencyNumbers += New-CsTeamsEmergencyNumber -EmergencyDialString $row.EmergencyDialString -EmergencyDialMask $row.EmergencyDialMask -OnlinePSTNUsage $row.OnlinePSTNUsage
            }
            New-CsTeamsEmergencyCallRoutingPolicy -Identity $row.Name -Description $row.Description -AllowEnhancedEmergencyServices (GetBool -onoff $row.DynamicEmergencyCalling) -EmergencyNumbers $emergencyNumbers   
        }
    }
    Write-Host ""
    Write-Host "Emergency call routing policies done!" -ForegroundColor Green
    Write-Host ""
    pause
}

function CreateCallingPolicies {
    foreach ($row in $excelData) {
        $parameters = @{
            Identity = $row.Name
            Description = $row.Description
            AllowPrivateCalling = (GetBool -onoff $row.AllowPrivateCalling)
            AllowWebPSTNCalling = (GetBool -onoff $row.AllowWebPSTNCalling)
            AllowSIPDevicesCalling = (GetBool -onoff $row.AllowSIPDevicesCalling)
            AllowVoicemail = $row.AllowVoicemail
            AllowCallGroups = (GetBool -onoff $row.AllowCallGroups)
            AllowDelegation = (GetBool -onoff $row.AllowDelegation)
            AllowCallForwardingToUser = (GetBool -onoff $row.AllowCallForwardingToUser)
            AllowCallForwardingToPhone = (GetBool -onoff $row.AllowCallForwardingToPhone)
            PreventTollBypass = (GetBool -onoff $row.PreventTollBypass)
            BusyOnBusyEnabledType = $row.BusyOnBusyEnabledType
            MusicOnHoldEnabledType = $row.MusicOnHoldEnabledType
            AllowCloudRecordingForCalls = (GetBool -onoff $row.AllowCloudRecordingForCalls)
            AllowTranscriptionForCalling = (GetBool -onoff $row.AllowTranscriptionForCalling)
            PopoutForIncomingPstnCalls = $row.PopoutForIncomingPstnCalls
            PopoutAppPathForIncomingPstnCalls = $row.PopoutAppPathForIncomingPstnCalls
            LiveCaptionsEnabledTypeForCalling = $row.LiveCaptionsEnabledTypeForCalling
            AutoAnswerEnabledType = $row.AutoAnswerEnabledType
            SpamFilteringEnabledType = $row.SpamFilteringEnabledType
            CallRecordingExpirationDays = $row.CallRecordingExpirationDays
            AllowCallRedirect = $row.AllowCallRedirect
            InboundPstnCallRoutingTreatment = $row.InboundPstnCallRoutingTreatment
            InboundFederatedCallRoutingTreatment = $row.InboundFederatedCallRoutingTreatment
        }

        $getCallingPolicy = Get-CsTeamsCallingPolicy -Identity $row.Name -ErrorAction SilentlyContinue

        if(![string]::IsNullOrEmpty($getCallingPolicy)) {
            Clear-Host
            Write-Host "Calling policy already exists!" -ForegroundColor Yellow
            Write-Host $row.Name -ForegroundColor Yellow
            $choose = Read-Host "Do you want to update the settings? [y/N]"
            if ($choose -eq "y") {
                Set-CsTeamsCallingPolicy @parameters
                Write-Host ""
                Write-Host "Calling policy updated successfully!" -ForegroundColor Green
                Write-Host ""
                pause
            }
        } else {
            New-CsTeamsCallingPolicy @parameters
        }
    }
    Write-Host ""
    Write-Host "Calling policies done!" -ForegroundColor Green
    Write-Host ""
    pause
}

function CreateCallHoldPolicies {
    foreach ($row in $excelData) {
        $getCallHoldPolicy = Get-CsTeamsCallHoldPolicy -Identity $row.Name -ErrorAction SilentlyContinue

        if(![string]::IsNullOrEmpty($getCallHoldPolicy)) {
            Clear-Host
            Write-Host "Call hold policy already exists!" -ForegroundColor Yellow
            Write-Host $row.Name -ForegroundColor Yellow
            $choose = Read-Host "Do you want to update the settings? [y/N]"
            if ($choose -eq "y") {
                $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
                    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
                    Filter = 'Sound Files (*.mp3;*.wma)|*.mp3;*.wma'
                }
                $null = $FileBrowser.ShowDialog()
                $content = Get-Content $FileBrowser.FileName -Encoding byte -ReadCount 0
                $audioFile = Import-CsOnlineAudioFile -FileName $FileBrowser.SafeFileName -Content $content
                Set-CsTeamsCallHoldPolicy -Identity $row.Name -Description $row.Description -AudioFileId $audioFile.Id
                Write-Host ""
                Write-Host "Call hold policy updated successfully!" -ForegroundColor Green
                Write-Host ""
                pause
            }
        } else {
            $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
                InitialDirectory = [Environment]::GetFolderPath('Desktop') 
                Filter = 'Sound Files (*.mp3;*.wma)|*.mp3;*.wma'
            }
            $null = $FileBrowser.ShowDialog()
            $content = Get-Content $FileBrowser.FileName -Encoding byte -ReadCount 0
            $audioFile = Import-CsOnlineAudioFile -FileName $FileBrowser.SafeFileName -Content $content
            New-CsTeamsCallHoldPolicy -Identity $row.Name -Description $row.Description -AudioFileId $audioFile.Id
        }
        
    }
    Write-Host ""
    Write-Host "Call hold policies done!" -ForegroundColor Green
    Write-Host ""
    pause
}

function CreateDialPlan {
    $dialPlanData = $excelData | Group-Object -Property DialPlan

    foreach ($data in $dialPlanData) {
        $getDialPlan = Get-CsTenantDialPlan -Identity $data.Name -ErrorAction SilentlyContinue

        $normRules = @()

        foreach ($row in $data.Group) {

            $id1 = "Global/" + $row.RuleName
            $nr = New-CsVoiceNormalizationRule -Identity $id1 -Description $row.RuleDescription -Pattern $row.Pattern -Translation $row.Translation -InMemory
            $normRules += $nr

        }

        if(![string]::IsNullOrEmpty($getDialPlan)) {
            Write-Host "Dial plan already exists!" -ForegroundColor Yellow
            Write-Host $data.Name -ForegroundColor Yellow
            $choose = Read-Host "Do you want to update the settings? [y/N]"
            if ($choose -eq "y") {
                Set-CsTenantDialPlan -Identity $data.Name -NormalizationRules $normRules
                Clear-Host
                Write-Host ""
                Write-Host "Dial plan updated successfully!" -ForegroundColor Green
                Write-Host ""
                pause
            }
        } else {
            New-CsTenantDialPlan -Identity $data.Name -NormalizationRules $normRules
        }

    }
    Clear-Host
    Write-Host ""
    Write-Host "Dial plans done!" -ForegroundColor Green
    Write-Host ""
    pause
}

function CreateUsers {
    $i=2
    [System.Collections.ArrayList]$newUsersArray = @{}
    
    foreach ($row in $excelData) {
        $newUsersArray.Add([pscustomobject]@{'ExcelRow'=$i;'UPN'=$row.UPN;'Number'=$row.Number;'Calling Policy'=$row.'Calling Policy';'Dial Plan'=$row.'Dial Plan';'Voice Routing Policy'=$row.'Voice Routing Policy';'Call Hold Policy'=$row.'Call Hold Policy';'Call Park Policy'=$row.'Call Park Policy';'Caller ID Policy'=$row.'Caller ID Policy';'Emergency Policy'=$row.'Emergency Policy';'Voice Mail'=$row.'Voice Mail';'Generated'=$row.Generated})
        $i++
    }

    $filteredUsers = $newUsersArray | Where-Object { $_.Generated -ne 'y'}

    # Batch Assign Policies to Users
    $callingPolicy_Users = $filteredUsers | Group-Object 'Calling Policy'
    $dialPlan_Users = $filteredUsers | Group-Object 'Dial Plan'
    $voiceRoutingPolicy_Users = $filteredUsers | Group-Object 'Voice Routing Policy'
    $callParkPolicy_Users = $filteredUsers | Group-Object 'Call Park Policy'
    $callerIdPolicy_Users = $filteredUsers | Group-Object 'Caller ID Policy'
    $emergencyPolicy_Users = $filteredUsers | Group-Object 'Emergency Policy'

    # TeamsCallingPolicy
    foreach ($data in $callingPolicy_Users) {
        $userIds = @()
        $policyName = $data.Name

        if([string]::IsNullOrEmpty($policyName)) { continue } 

        foreach ($user in $data.Group) {
            $userIds += $user.UPN
        }
        New-CsBatchPolicyAssignmentOperation -PolicyType TeamsCallingPolicy -PolicyName $policyName -Identity $userIds -OperationName "Batch assign $policyName"

    }

    # TenantDialPlan
    foreach ($data in $dialPlan_Users) {
        $userIds = @()
        $policyName = $data.Name

        if([string]::IsNullOrEmpty($policyName)) { continue }

        foreach ($user in $data.Group) {
            $userIds += $user.UPN
        }
        New-CsBatchPolicyAssignmentOperation -PolicyType TenantDialPlan -PolicyName $policyName -Identity $userIds -OperationName "Batch assign $policyName"

    }

    # OnlineVoiceRoutingPolicy
    foreach ($data in $voiceRoutingPolicy_Users) {
        $userIds = @()
        $policyName = $data.Name

        if([string]::IsNullOrEmpty($policyName)) { continue }

        foreach ($user in $data.Group) {
            $userIds += $user.UPN
        }
        New-CsBatchPolicyAssignmentOperation -PolicyType OnlineVoiceRoutingPolicy -PolicyName $policyName -Identity $userIds -OperationName "Batch assign $policyName"

    }

    # TeamsCallParkPolicy
    foreach ($data in $callParkPolicy_Users) {
        $userIds = @()
        $policyName = $data.Name

        if([string]::IsNullOrEmpty($policyName)) { continue }

        foreach ($user in $data.Group) {
            $userIds += $user.UPN
        }
        New-CsBatchPolicyAssignmentOperation -PolicyType TeamsCallParkPolicy -PolicyName $policyName -Identity $userIds -OperationName "Batch assign $policyName"

    }

    # CallingLineIdentity
    foreach ($data in $callerIdPolicy_Users) {
        $userIds = @()
        $policyName = $data.Name

        if([string]::IsNullOrEmpty($policyName)) { continue }

        foreach ($user in $data.Group) {
            $userIds += $user.UPN
        }
        New-CsBatchPolicyAssignmentOperation -PolicyType CallingLineIdentity -PolicyName $policyName -Identity $userIds -OperationName "Batch assign $policyName"

    }

    # TeamsEmergencyCallRoutingPolicy
    foreach ($data in $emergencyPolicy_Users) {
        $userIds = @()
        $policyName = $data.Name

        if([string]::IsNullOrEmpty($policyName)) { continue }

        foreach ($user in $data.Group) {
            $userIds += $user.UPN
        }
        New-CsBatchPolicyAssignmentOperation -PolicyType TeamsEmergencyCallRoutingPolicy -PolicyName $policyName -Identity $userIds -OperationName "Batch assign $policyName"

    }

    foreach ($row in $filteredUsers) {
        $upn = $row.UPN
        $number = $row.Number
        $callHoldPolicy = $row.'Call Hold Policy'
        $voiceMail = $row.'Voice Mail'
        $excelRow = $row.ExcelRow

        if(![string]::IsNullOrEmpty($number)) {
            try {
                Set-CsPhoneNumberAssignment -Identity $upn -PhoneNumber $number -PhoneNumberType DirectRouting -ErrorAction SilentlyContinue
            } catch {
                write-host "Cannot add number to $upn" -ForegroundColor Yellow
                LogWrite "Cannot add number to $upn"
                LogWrite $error[0].Exception
            }
        } else {
             try {
                Set-CsPhoneNumberAssignment -Identity $upn -EnterpriseVoiceEnabled $true -ErrorAction SilentlyContinue
            }
            catch {
                write-host "Cannot enable Enterprise Voice on $upn" -ForegroundColor Yellow
                LogWrite "Cannot enable Enterprise Voice on $upn"
                LogWrite $error[0].Exception
            }
        }

        if(![string]::IsNullOrEmpty($callHoldPolicy)) {
            Grant-CsTeamsCallHoldPolicy -Identity $upn -PolicyName $callHoldPolicy -ErrorAction SilentlyContinue
        }

        if(![string]::IsNullOrEmpty($voiceMail)) {
            Set-CsOnlineVoicemailUserSettings -Identity $upn -PromptLanguage $voiceMail -ErrorAction SilentlyContinue
        }

    }

}

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

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'SpreadSheet (*.xlsx)|*.xlsx'
}
Write-Host "Please choose the runbook excel file."
$null = $FileBrowser.ShowDialog()

if([string]::IsNullOrEmpty($FileBrowser.FileName)) {
    break
}

Clear-Host
Write-Host "Please enter your Tenant ID (can be left empty if you have only one)" -ForegroundColor Green
$tenantID = Read-Host -Prompt "Tenant ID"
if(![string]::IsNullOrEmpty($tenantID)) {
    Connect-MicrosoftTeams -TenantId $tenantID
} else {
   Connect-MicrosoftTeams 
}


$excelpath = $FileBrowser.FileName

function ShowMenu {
    Clear-Host
    Write-Host "==================================="
    Write-Host ""
    Write-Host "     MS Teams Workbook Script"
    Write-Host ""
    Write-Host "==================================="
    Write-Host ""
    Write-Host "[1] SBC"
    Write-Host "[2] PSTN Usage"
    Write-Host "[3] Voice Routes"
    Write-Host "[4] Voice Routing Policies"
    Write-Host "[5] Emergency Call Routing Policies"
    Write-Host "[6] Calling Policies"
    Write-Host "[7] Call Hold Policies"
    Write-Host "[8] Dial Plan"
    Write-Host "[9] Import Users"
    Write-Host "-----------------------------------"
    Write-Host "[Q] Quit"
    Write-Host "-----------------------------------"
}

do {
    ShowMenu
    $choose = Read-Host "Please choose an option"

    switch ($choose) {
        "1" {
            $workSheetName = "SBC"
            $excelData = ReadExcelFile
            CreateSBC
        }
        "2" {
            $workSheetName = "PSTN Usage"
            $excelData = ReadExcelFile
            CreatePSTNUsages
        }
        "3" {
            $workSheetName = "Voice Routes"
            $excelData = ReadExcelFile
            CreateVoiceRoutes
        }
        "4" {
            $workSheetName = "Voice Routing Policy"
            $excelData = ReadExcelFile
            CreateVoiceRoutingPolicies
        }
        "5" {
            $workSheetName = "Emergency Call Routing Policy"
            $excelData = ReadExcelFile
            CreateEmergencyCallRoutingPolicies
        }
        "6" {
            $workSheetName = "Calling Policy"
            $excelData = ReadExcelFile
            CreateCallingPolicies
        }
        "7" {
            $workSheetName = "Call Hold Policy"
            $excelData = ReadExcelFile
            CreateCallHoldPolicies
        }
        "8" {
            $workSheetName = "Dial Plan"
            $excelData = ReadExcelFile
            CreateDialPlan
        }
        "9" {
            $workSheetName = "Users"
            $excelData = ReadExcelFile
            CreateUsers
        }
    }

} until ($choose -eq 'q')

Clear-Host