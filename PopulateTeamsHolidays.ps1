# PopulateTeamsHolidays.PS1
# Update/Create the Teams holiday schedule with new events
#
# Version 1.1.0 (Build 1.1.0-2023-01-31)
# 
# MG & AT
#
#########################################################################################
#
# JUST RUN THE SCRIPT!
#
#########################################################################################
#                            DO NOT EDIT BELOW THESE LINES!                             #
#########################################################################################
#
# For supported countries, please visit https://www.openholidaysapi.org/
#
#
$country = ""
$province = ""
$schedulename = ""
$tenantID = ""
#
# Fill in Local path of CSV file for offline usage or not currently supported countries
#
$localFile = ""
$scheduleid = ""
#
$startDate = Get-Date -Format "yyyy-MM-dd"
$endDate = Get-Date((Get-Date).AddYears(1)) -format "yyyy-MM-dd"
#
#########################################################################################

function Get-Countries {
    $Countries = @()
    $countries_uri = "https://openholidaysapi.org/Countries?languageIsoCode=EN"

    $headers = @{
      'Accept'       = 'application/json'
      'Content-Type'  = 'application/json'
    }

    $Countries = Invoke-RestMethod -Uri $countries_uri -Method GET -Headers $headers

    return $Countries
}

function Get-Subdivisions {
    param(
        [Parameter (Mandatory = $false)] [String]$IsoCode
    )
    $Subdivisions = @()
    $subdivisions_uri = "https://openholidaysapi.org/Subdivisions?countryIsoCode=" + $IsoCode + "&languageIsoCode=EN"

    $headers = @{
      'Accept'       = 'application/json'
      'Content-Type'  = 'application/json'
    }

    $Subdivisions = Invoke-RestMethod -Uri $subdivisions_uri -Method GET -Headers $headers

    return $Subdivisions
}

function Get-Schedules {
    [array]$getAllScheduler = Get-CsOnlineSchedule

    return $getAllScheduler
}

Clear-Host

Write-Host "Please enter your Tenant ID (can be left empty if you have only one)" -ForegroundColor Green
$tenantID= Read-Host -Prompt "Tenant ID"
If($tenantID.Length -gt 10) {
        Connect-MicrosoftTeams -TenantId $tenantID
    } else {
        Connect-MicrosoftTeams
    }
Clear-Host
Write-Host "Please choose if you want to create a new Holiday or load from MS Teams" -ForegroundColor Green
Write-Host "1. Create New (Default)"
Write-Host "2. Load"
$createOrLoad = Read-Host -Prompt "Choose"
If($createOrLoad -eq 2) {
    Clear-Host
    Write-Host "Please choose one of the existing Holidays from MS Teams" -ForegroundColor Green
    $i=1
    $schedulerList = Get-Schedules
    ForEach($schedulerEntry in $schedulerList) {
    If($schedulerEntry.Name -eq $null) {
        continue
        }
         Write-Host("["+ $i++ + "] " + $schedulerEntry.Name + " (" + $schedulerEntry.Id + ")")
    }
    $scheduler = Read-Host -Prompt "Choose"
    If($scheduler -eq "") {
        Clear-Host
        Write-Host "Please enter a new Name for the Holiday Table" -ForegroundColor Green
        $schedulename = Read-Host -Prompt "Name"
    } else {
        $scheduleid = $schedulerList[$scheduler-1].Id
    }
} else {
    Clear-Host
    Write-Host "Please enter a new Name for the Holiday Table" -ForegroundColor Green
    $schedulename = Read-Host -Prompt "Name"
}
Clear-Host
Write-Host "Please choose if you want to create a new Holiday or load from MS Teams" -ForegroundColor Green
Write-Host "1. Local File"
Write-Host "2. Online (Default)"
$source = Read-Host -Prompt "Holiday Source"

If($source -eq 1) {
    Clear-Host
    Write-Host "Please enter the file including the path. (e.g. C:\Temp\Holiday.CSV)" -ForegroundColor Green
    $localFile = Read-Host -Prompt "File"
} else {
    Clear-Host
    Write-Host "Please choose the country and enter the 2-letter country code" -ForegroundColor Green
    ForEach($countryEntry in Get-Countries) {
         Write-Host("["+$countryEntry.isoCode+"] " + $countryEntry.name.text)
    }

    $country = Read-Host "Country Code"

    Clear-Host
    Write-Host "Please choose the state and enter the state code." -ForegroundColor Green
    Write-Host "Leave empty for all holidays from the choosen country." -ForegroundColor Green

    ForEach($provinceEntry in (Get-Subdivisions -IsoCode $country)) {
         Write-Host("["+$provinceEntry.isoCode+"] " + $provinceEntry.name.text)
    }

    $province = Read-Host "State Code"

}

function Get-PublicHolidays {
    $Holidays = @()
    $holiday_uri = "https://openholidaysapi.org/PublicHolidays?countryIsoCode=" + $country + "&languageIsoCode=DE&validFrom=" + $startDate + "&validTo=" + $endDate + "&subdivisionCode=" + $province

    $headers = @{
      'Accept'       = 'application/json'
      'Content-Type'  = 'application/json'
    }

    $Holidays = Invoke-RestMethod -Uri $holiday_uri -Method GET -Headers $headers

    return $Holidays
}

# Read in public holidays file
If($localFile -eq "") {
    $PublicHolidays = Get-PublicHolidays
} else {
    $PublicHolidays = Import-Csv $localFile
}

Clear-Host
Write-Host "Creating Holiday Table" -ForegroundColor Green
Write-Host ""
# Process each event from the holidays file
$i=0

# create array of holidays according lenght of publicholidays
$HolidayDateRange= [Microsoft.Rtc.Management.Hosted.Online.Models.DateTimerange[]]::new($PublicHolidays.Count)

# Process each event from the holidays
ForEach ($PublicHoliday in $PublicHolidays) {
   If ($localFile -eq "") {
      $publicHolidayName = $PublicHoliday.name.text
      $publicHolidayDate = $PublicHoliday.startDate
   } else {
      $publicHolidayName = $PublicHoliday.Holiday
      $publicHolidayDate = $PublicHoliday.Date
   }
      Write-Host ("Processing {0} on {1}" -f $publicHolidayName, $publicHolidayDate)
      $dd = Get-Date($publicHolidayDate) -format dd
      $mm = Get-Date($publicHolidayDate) -format MM
      $yyyy = Get-Date($publicHolidayDate) -format yyyy
      $date = $dd + "/" + $mm +"/"+ $yyyy
      $holiday = New-CsOnlineDateTimeRange -Start $Date
      $HolidayDateRange[$i++] = $holiday
    
}

break;

# Change or Create scheduler with new Holiday dates
If($scheduleid -eq "") {
    Write-Host ""
    Write-Host "Upload new Holiday Table" -ForegroundColor Green
    Write-Host ""
    New-CsOnlineSchedule -Name $schedulename -FixedSchedule -DateTimeRanges @($HolidayDateRange)
}
else {
    Write-Host ""
    Write-Host ("Updating Holiday Table (" + $scheduleid + ")") -ForegroundColor Green
    Write-Host ""
    $HolidayID = Get-CsOnlineSchedule -Id $scheduleid
    $HolidayID.FixedSchedule.DateTimeRanges = @($HolidayDateRange)
    Set-CsOnlineSchedule -Instance $HolidayID
}

Write-Host ""
Write-Host "ALL DONE!" -ForegroundColor Green
Write-Host ""