<#
    .SYNOPSIS
        Searches through a given path and scans all Excel files found in there.
        In the Excel files it searches for the name of the logged in user in column one. If the user name is found, go through all the days and search for shifts
    
    .DESCRIPTION


    .INPUTS
        Parameter "MonthToSync". Defines which month to sync, to reduce the risk in case of an error
        Parameter "ExcelFilePath". If a path is given, script will only run on that file. If empty, script will run through all the files in the path defined further below
        Parameter "LogVerbose". If 1 INFO Logs will be written to file. If 0 only warnings and worse will be written
    
    .OUTPUTS
        Log file stored in C:\Temp\<name>.log (Change LogFile Variabel if you want a different location)
  
    .NOTES
        File Name: Sync-ExcelOutlook.ps1
        Author: Philip Schwander
        
        Changelog:
            11.04.18 Philip Schwander - Creation of script
            16.04.18 Philip Schwander - Seperate tags for morning and afternoon added
            17.04.18 Philip Schwander - Added functionality to protect manully modified Outlook entries by adding ## to the subject.
                                        Non protected entries will be deleted and recreated or just deleted if they were deleted in Excel.
                                        Logging can now be switched to non-verbose for small log files
            18.04.18 Philip Schwander - Added support for all day events
            23.07.18 Philip Schwander - Changed code to work with new spreadsheet design (Cell for year has moved, some more exeptions, client list has moved)
            24.07.18 Philip Schwander - Appointments weren't deleted anymore. Line 253 removed hashtag to include the delete statement
            26.07.18 Philip Schwander - Changed Exception Check to exclude all Shift Strings that start with an exclamation mark (!) (like !./!x and !WE)
            
        © Upgreat AG
    .COMPONENT
        
    .LINK
        
#>

param (
    [string]$MonthToSync = '',
    [string]$ExcelFilePath = '',
    [int]$LogVerbose = 0
)

# Enum to be used for the busy state in appointments
Add-Type -TypeDefinition @"
   public enum BusyOptions
   {
      Free,
      Tentative,
      Busy,
      OutOfOffice
   }
"@

# Enum to define morning and afternoon
Add-Type -TypeDefinition @"
   public enum TimeOfDay
   {
      Undefined,
      AM,
      PM,
      AllDay
   }
"@

function Set-Culture([System.Globalization.CultureInfo] $culture)
{
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = $culture
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $culture
}

#Set-Culture en-US

# Define months
$MonthsToNumbers = @{}
$MonthsToNumbers.Add('Januar',1)
$MonthsToNumbers.Add('Februar',2)
$MonthsToNumbers.Add('März',3)
$MonthsToNumbers.Add('April',4)
$MonthsToNumbers.Add('Mai',5)
$MonthsToNumbers.Add('Juni',6)
$MonthsToNumbers.Add('Juli',7)
$MonthsToNumbers.Add('August',8)
$MonthsToNumbers.Add('September',9)
$MonthsToNumbers.Add('Oktober',10)
$MonthsToNumbers.Add('November',11)
$MonthsToNumbers.Add('Dezember',12)

# Define client list
$ClientDict = @{}


# Logging function
Function Write-Log {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$False)]
    [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
    [String]
    $Level = "INFO",

    [Parameter(Mandatory=$True)]
    [string]
    $Message,

    [Parameter(Mandatory=$False)]
    [string]
    $logfile
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $Level $Message"
    If($logfile) {
        # If LogVerbose is true, write all info to file. If not only write WARN and worse
        if ($LogVerbose -eq 0) {
            if(($Level -eq "INFO") -or ($Level -eq "DEBUG")) {
                
            } else {
                Add-Content $logfile -Value $Line    
            }
        } else {
            Add-Content $logfile -Value $Line
        }
    }
    Else {
        Write-Output $Line
    }
}

#$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
#Write-Host $StopWatch.ElapsedMilliseconds

# Log variables
$LogStart = [System.DateTime]::Now
$LogDate = (Get-Date -Format "MM-dd-yyyy")
$LogFile = "C:\Temp\Sync $LogDate.log"


# Path to where the planning Excel files are located
$ExcelFileDirectory = "\\ic.up-great.ch@ssl\Organization\TeamSites\ServiceOperations\Documents\01_Planning"


# Create Regex for "_dddd_" like _2018_
$YearRegex = "[\d{1,5}]{4}"

# Create Regex for '*##*'
$HashtagRegex = '*##*'


# Start Column -1 so day numbers work out properly
$startRow,$col=1,3
$lastRow = 100

# Row and column for clients
$ClientRow = 7
$ClientCol = 1
$ClientLastRow = 14
$ClientLastCol = 1


$Culture = New-Object system.globalization.cultureinfo("de-DE")


$StartTime = Get-Date -Hour 0 -Minute 0
$EndTime = Get-Date -Hour 23 -Minute 59



# Variable for sync flag. This is set in the billing info attribute of the appointment to identify entries created by this script
$SyncManagedFlag = 'ManagedBySync'

# Variable to hold the time of day
[TimeOfDay]$ToD = [TimeOfDay]::Undefined

# Busy status for appointment 0 = Free, 1 = Tentative, 2 = Busy, 3 = Out of Office
[BusyOptions]$BusyStatus = [BusyOptions]::Free

# All day event variable
$AllDayEvent = $False



# Create Outlook Application Item
$olFolderCalendar = 9
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace('MAPI')


# Create appointment variable
$Appointments = $ns.GetDefaultFolder($olFolderCalendar).Items
$Appointments.Sort("[Start]")
$Appointments.IncludeRecurrences = $false

# Ressourcen Plan
$ShiftAddFlag = " (RP)"

# Categories variable
$AppointmentCategories = ''


# Create Excel Object
$excel=new-object -com excel.application


# Get logged in user, cut the domain off
$UserName = whoami
$UserName = $UserName.Split("\")[1]
$NoEntryForUser = $False

# Write log start info
Write-Log -Message "`r`n" -logfile $LogFile
Write-Log -Level INFO -Message "Start Sync $LogStart" -logfile $LogFile
Write-Log -Level INFO -Message "Searching for user $UserName" -logfile $LogFile



if ( $MonthToSync -eq '') {
    Write-Log -Level INFO -Message ("Parameter MonthToSync is empty string, nothing will be synced") -logfile $LogFile
} else {
    Write-Log -Level INFO -Message ("Syncing month: " + $MonthToSync) -logfile $LogFile
}

#### TEST #####
#$ExcelFilePath = "\\up-great.local\shares\benutzer\pschwander\Documents\ITOPS2_Ressourcenplan_2018_TestCopy.xlsm"
#$MonthToSync = "April"

# Check if path has been given as parameter and if yes set it as the only file to go through
if ($ExcelFilePath -eq '') {
    $FileList = Get-ChildItem $ExcelFileDirectory -Filter *.xlsx
} else {
    $FileList = $ExcelFilePath
}

# Iterate through the file list
Foreach($Filepath in $FileList) {
    # Get year from filename. Find _yyyy_, remove leading and trailing underscore and then convert to date time [year]
    #$found = ($Filepath -match $YearRegex)
    #$PlanningYear = $matches[0].Substring(1,$matches[0].Length -2)
    #$PlanningYear = [datetime]::ParseExact($PlanningYear, "yyyy", $null);

    # Get Workbook from file
    $WorkBook=$excel.workbooks.open($Filepath)

    ### Create client dictionary and get year
    $WorkSheet=$WorkBook.Sheets.Item(13)
    $PlanningYear = $WorkSheet.Cells.Item(3,2).Value2
    $PlanningYear = [datetime]::ParseExact($PlanningYear, "yyyy", $null);

    # Read all clients from the data sheet in Excel (PCH eq Porsche etc.) and build a dictionary
    for ($yy=$ClientRow; $yy -le $ClientLastRow; $yy++) {
        for ($ww=$ClientCol; $ww -le $ClientLastCol; $ww+=3) {
            $Client = $WorkSheet.Cells.Item($yy,$ww).Value2
            $ClientFullName = $WorkSheet.Cells.Item($yy,$ww+1).Value2
            if (![string]::IsNullOrEmpty($Client)) {
                $ClientDict.Add($Client, $ClientFullName)
            }
        }
    }

    
    $WorkSheet=$WorkBook.Sheets.Item(1)
    #$YearCellValue = $WorkSheet.Cells.Item(3,2).Value2
    #$found = ($YearCellValue -match $YearRegex)
    #$PlanningYear = $YearCellValue
    #$PlanningYear = [datetime]::ParseExact($PlanningYear, "yyyy", $null);

    Write-Log -Level INFO -Message ("Sync Workbook " + $_.FullName) -logfile $LogFile
    Write-Log -Level INFO -Message "Working on sync for year $PlanningYear" -logfile $LogFile

    # Iterate through all the sheets in the workbook
    for ($i=1; $i -le $WorkBook.sheets.count; $i++) {
        
        # Check if we're in the month we want to sync
        if ($MonthsToNumbers[$MonthToSync] -ne $i) {
            continue
        }
        # Get the sheet from the workbook
        $WorkSheet=$WorkBook.Sheets.Item($i)


        # Find the logged on users line
        for ($xx=$startRow; $xx -le $lastRow; $xx++) {
            #Write-Host $WorkSheet.Cells.Item($xx,1).Value2
            if ($UserName -eq $WorkSheet.Cells.Item($xx,1).Value2) {
                Write-Log -Level INFO -Message "User found $UserName" -logfile $LogFile
                Write-Log -Level INFO -Message "User start row $xx" -logfile $LogFile
                $firstRow = $xx
                $NoEntryForUser = $False
                break
            } else {
                $NoEntryForUser = $True
            }
        }

        # Exit this worksheet because user isn't found
        if ($NoEntryForUser -eq $True) {
            Write-Log -Level ERROR -Message "No entries found for user $UserName in this month work sheet" -logfile $LogFile
            Write-Log -Level INFO -Message "Exiting work sheet" -logfile $LogFile
            $NoEntryForUser = $False    
            continue
        }
     
        Write-Log -Level INFO -Message "Checking for shifts in work sheet nr. $i" -logfile $LogFile

        # Iterate through all the days
        for ($j=1; $j -le ([datetime]::DaysInMonth($PlanningYear.Year,$i)); $j++) {

            # Read all lines belonging to the user (morning and afternoon and on-call)
            for ($l=0; $l -le 2; $l++) {
                switch($l)
                {
                    0 {
                        $ToD = [TimeOfDay]::AM
                        $StartTime = Get-Date -Hour 7 -Minute 0
                        $EndTime = Get-Date -Hour 12 -Minute 59
                    }
                    1 {
                        $ToD = [TimeOfDay]::PM
                        $StartTime = Get-Date -Hour 12 -Minute 0
                        $EndTime = Get-Date -Hour 18 -Minute 0
                    }
                    2 {
                        $ToD = [TimeOfDay]::AllDay
                    }
                }
     
                # Get the value from the calculated Excel cell and store it in variable Shift
                $Shift=$WorkSheet.Cells.Item(($firstRow + $l),($col + $j)).Value2

                # Check if string is not empty
                if(![string]::IsNullOrEmpty($Shift) -and (!$Shift.StartsWith("!"))) {
                    Write-Log -Level INFO -Message "Found a shift $Shift" -logfile $LogFile

                    # Replace abbreviated string with full string according to Excel workbook and set busy state
                    if ( $Shift -eq 'H' ) { $Shift = 'HomeOffice'; $BusyStatus = [BusyOptions]::OutOfOffice }
                    elseif ( $Shift -eq 'S' ) { $Shift = 'Schulung'; $BusyStatus = [BusyOptions]::OutOfOffice}
                    elseif ( $Shift -eq 'R') { $Shift = 'Ressort'; $BusyStatus = [BusyOptions]::Busy }
                    elseif ( $Shift -eq 'P' ) { $Shift = 'Projekt'; $BusyStatus = [BusyOptions]::Busy }
                    elseif ( $Shift -eq 'F' ) { $Shift = 'Ferien'; $BusyStatus = [BusyOptions]::OutOfOffice  }
                    elseif ( $Shift -eq 'KO' ) { $Shift = 'Kompensation'; $BusyStatus = [BusyOptions]::OutOfOffice}
                    elseif ( $Shift -eq 'K' ) { $Shift = 'Krank'; $BusyStatus = [BusyOptions]::OutOfOffice }
                    elseif ( $Shift -eq 'M' ) { $Shift = 'Militär'; $BusyStatus = [BusyOptions]::OutOfOffice }
                    elseif ( $Shift -eq 'A' ) { $Shift = 'Andere'; $BusyStatus = [BusyOptions]::Busy }
                    else {
                        # All client on-premise visites are handled here and marked as out of office
                        if(![string]::IsNullOrEmpty($ClientDict[$Shift])) {
                            $Shift = $ClientDict[$Shift]
                            $BusyStatus = [BusyOptions]::OutOfOffice
                        }
                    }

                    # Map times to shift and set categories and busy state
                    if ($Shift.Equals('SD1') -or $Shift.Equals('SV1')) {
                        $AppointmentCategories = ""
                        $StartTime = Get-Date -Hour 7 -Minute 0
                        $EndTime = Get-Date -Hour 16 -Minute 0
                        $AllDayEvent = $False
                        $BusyStatus = [BusyOptions]::Free
                    } elseif ($Shift.Equals('SD2') -or $Shift.Equals('SV2')) {
                        $AppointmentCategories = ""
                        $StartTime = Get-Date -Hour 8 -Minute 0
                        $EndTime = Get-Date -Hour 17 -Minute 0
                        $AllDayEvent = $False
                        $BusyStatus = [BusyOptions]::Free
                    } elseif ($Shift.Equals('SD3') -or $Shift.Equals('SV3')) {
                        $AppointmentCategories = ""
                        $StartTime = Get-Date -Hour 9 -Minute 0
                        $EndTime = Get-Date -Hour 18 -Minute 0
                        $AllDayEvent = $False
                        $BusyStatus = [BusyOptions]::Free
                    } elseif ($Shift.Equals('P1')) {
                        $Shift = 'Pikett UPG'
                        $AppointmentCategories = "P1"
                        $StartTime = Get-Date -Hour 6 -Minute 0
                        $EndTime = Get-Date -Hour 6 -Minute 5
                        $AllDayEvent = $True
                        $BusyStatus = [BusyOptions]::Free
                    } elseif ($Shift.Equals('P2')) {
                        $Shift = 'Pikett PCH'
                        $AppointmentCategories = "P2"
                        $StartTime = Get-Date -Hour 6 -Minute 0
                        $EndTime = Get-Date -Hour 6 -Minute 5
                        $AllDayEvent = $True
                        $BusyStatus = [BusyOptions]::Free
                    } elseif ($Shift.Equals('P3')) {
                        $Shift = 'Pikett DC'
                        $AppointmentCategories = "P3"
                        $StartTime = Get-Date -Hour 6 -Minute 0
                        $EndTime = Get-Date -Hour 6 -Minute 5
                        $AllDayEvent = $True
                        $BusyStatus = [BusyOptions]::Free
                    } else {
                        $AppointmentCategories = ''
                        $AllDayEvent = $False
                    }

                    # Create flag
                    $CreateNewEntry = $true

                    # Treat all day events different from the rest
                    if ($AllDayEvent) {
                        # Prepare Filter for search,
                        $Start = ((Get-Date -Year $PlanningYear.Year -Month $i -Day $j).AddDays(1)).ToShortDateString() + " 00:00"
                        $End = ((Get-Date -Year $PlanningYear.Year -Month $i -Day $j).AddDays(1)).ToShortDateString() + " 23:59"
                        $Filter = "[MessageClass]='IPM.Appointment' AND [AllDayEvent] = 'True' AND [Start] > '$Start' AND [End] < '$End'"

                        # Iterate through Appointments to see if current Excel entry already exists in Outlook and skip if that's the case
                        foreach ($Appointment in $Appointments.Restrict($Filter)) {
                            # Check if appointment has the proper tag
                            if ($Appointment.Subject -eq ($Shift + $ShiftAddFlag)) {
                                $CreateNewEntry = $False
                            }
                        }
                    
                    } else {

                        # Prepare Filter for search
                        #$Start = (Get-Date -Year $PlanningYear.Year -Month $i -Day $j).ToShortDateString() + " 00:00"
                        #$End = (Get-Date -Year $PlanningYear.Year -Month $i -Day $j).ToShortDateString() + " 23:59"
                        #$Filter = "[MessageClass]='IPM.Appointment' AND [Start] > '$Start' AND [End] < '$End'"
                        $Start = (Get-Date -Year $PlanningYear.Year -Month $i -Day $j)
                        $End = (Get-Date -Year $PlanningYear.Year -Month $i -Day $j)
                        $Start = $Start.Date.AddHours(0)
                        $End = $End.Date.AddHours(23)
						$End = $End.AddMinutes(59)
                        $Filter = "[MessageClass]='IPM.Appointment' AND [Start] > '"+(Get-Date($Start) -Format g)+"' AND [End] < '"+(Get-Date($End) -Format g)+"'"
                        #$Filter = "[MessageClass]='IPM.Appointment' AND [Start] > '"+(Get-Date($Start))+"' AND [End] < '"+(Get-Date($End))+"'"
                        $AppointmentStartTime = $Start.Date.AddHours($StartTime.Hour)
                        $AppointmentEndTime = $End.Date.AddHours($EndTime.Hour)
                        Write-Log -Level INFO -Message ("Filter: " + $Filter) -logfile $LogFile
                        # Iterate through Appointments to see if current Excel entry already exists in Outlook and skip if that's the case
                        foreach ($Appointment in $Appointments.Restrict($Filter)) {
                            Write-Log -Level INFO -Message ("Appointment Debug " + $Appointment.Start + " " + $Appointment.End) -logfile $LogFile
                            if(($Appointment.Start -eq $AppointmentStartTime) -and ($Appointment.End -eq $AppointmentEndTime)) {
                                Write-Log -Level INFO -Message ("Found Appointment for $i/$j/" + $PlanningYear.Year + " with matching date, subject is " + $Appointment.Subject + ". Possible modifictaion needed") -logfile $LogFile
                                if($Appointment.Subject -ne ($Shift + $ShiftAddFlag)) {
                                    Write-Log -Level INFO -Message ("Appointment for $i/$j/" + $PlanningYear.Year + " with Subject " + $Appointment.Subject + " instead of " + $Shift + " needs to be recreated") -logfile $LogFile
                                    $CreateNewEntry = $true
                                } else {
                                    $CreateNewEntry = $false
                                    break
                                }
                            } else {
                                $CreateNewEntry = $true
                            }
                        }
                    }
                    
                    $Start = (Get-Date -Year $PlanningYear.Year -Month $i -Day $j)
                    $End = (Get-Date -Year $PlanningYear.Year -Month $i -Day $j)
                    $AppointmentStartTime = $Start.Date.AddHours($StartTime.Hour)
                    $AppointmentEndTime = $End.Date.AddHours($EndTime.Hour)
                    
                    # Check if any modified preexisting regular entries exist
                    foreach ($Appointment in $Appointments.Restrict($Filter)) {
                        # Check if an appointment exists with a morning or afternoon tag
                        if($Appointment.BillingInformation -eq ($SyncManagedFlag + $ToD)) {
							Write-Log -Level INFO -Message ("Appointment start time: " + $Appointment.Start + " Calculated start time: " + $AppointmentStartTime) -logfile $LogFile
							Write-Log -Level INFO -Message ("Appointment end time: " + $Appointment.End + " Calculated end time: " + $AppointmentEndTime) -logfile $LogFile
                            # Check if it's not fitting a correct entry
                            if((($Appointment.Start -ne $AppointmentStartTime) -or ($Appointment.End -ne $AppointmentEndTime)) -and -not $Appointment.AllDayEvent) {
                                # Check if it was modified on purpose (needs to have two consecutive hastags (##) in the subject
                                if ($Appointment.Subject -like $HashtagRegex) {
                                    # Modified on purpose - leave it
                                    Write-Log -Level INFO -Message ("Modified entry found $i/$j/" + $PlanningYear.Year + " with Subject " + $Appointment.Subject + " containing ##. No need to create a new one or modifying it") -logfile $LogFile
                                    $CreateNewEntry = $false
                                } else {
                                    # Modified accidentally - delete it
                                    Write-Log -Level INFO -Message ("Modified entry found $i/$j/" + $PlanningYear.Year + " with Subject " + $Appointment.Subject + " deleting it, a new one will be created ...") -logfile $LogFile
                                    $Appointment.Delete()
                                }
                            } elseif ($Appointment.Subject -ne ($Shift + $ShiftAddFlag)) {
                                # Check if it was modified on purpose (needs to have two consecutive hastags (##) in the subject
                                if ($Appointment.Subject -like $HashtagRegex) {
                                    # Modified on purpose - leave it
                                    Write-Log -Level INFO -Message ("Modified entry found $i/$j/" + $PlanningYear.Year + " with Subject " + $Appointment.Subject + " containing ##. No need to create a new one or modifying it") -logfile $LogFile
                                    $CreateNewEntry = $false
                                } else {
                                    # Modified accidentally or shift changed - delete it
                                    Write-Log -Level INFO -Message ("Modified entry found $i/$j/" + $PlanningYear.Year + " with Subject " + $Appointment.Subject + " deleting it, a new one will be created") -logfile $LogFile
                                    $Appointment.Delete()
                                }
                            }
                        }
                    }                    

                    # Create new calendar entry
                    if($CreateNewEntry) {
                        $Start = (Get-Date -Year $PlanningYear.Year -Month $i -Day $j)
                        $End = (Get-Date -Year $PlanningYear.Year -Month $i -Day $j)
                        $AppointmentStartTime = $Start.Date.AddHours($StartTime.Hour)
						#$AppointmentStartTime = $AppointmentStartTime.Date.AddMinutes(0)
						#$AppointmentStartTime = $AppointmentStartTime.Date.AddSeconds(0)
                        $AppointmentEndTime = $End.Date.AddHours($EndTime.Hour)
						#$AppointmentEndTime = $AppointmentEndTime.Date.AddMinutes(0)
						#$AppointmentEndTime = $AppointmentEndTime.Date.AddSeconds(0)
                        Write-Log -Level INFO -Message ("Non existing entry found $i/$j/" + $PlanningYear.Year + " with Subject " + $Shift + " creating it") -logfile $LogFile
                        Write-Log -Level INFO -Message ("Creating new entry with " + $AppointmentStartTime + " " + $AppointmentEndTime) -logfile $LogFile
                        Write-Log -Level INFO -Message ("Creating new entry with " + $StartTime + " " + $EndTime) -logfile $LogFile
                        $ol = New-Object -ComObject Outlook.Application
                        $meeting = $ol.CreateItem('olAppointmentItem')
                        $meeting.Subject = $Shift + $ShiftAddFlag
                        $meeting.BillingInformation = $SyncManagedFlag + $ToD
                        $meeting.Location = ''
                        $meeting.ReminderSet = $false
                        $meeting.Importance = 1
                        #$meeting.MeetingStatus = [Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeeting
                        #$meeting.ReminderMinutesBeforeStart = 15
                        $meeting.Start = $AppointmentStartTime
                        $meeting.End = $AppointmentEndTime
                        $meeting.Categories = $AppointmentCategories
                        $meeting.BusyStatus = $BusyStatus
                        $meeting.AllDayEvent = $AllDayEvent
                        $meeting.Save()
                    }
                } else {
                    # Prepare Filter in order to only scan a limited number of entries
                    #$Start = ((Get-Date -Year $PlanningYear.Year -Month $i -Day $j).AddDays(1)).ToShortDateString() + " 00:00"
                    #$End = ((Get-Date -Year $PlanningYear.Year -Month $i -Day $j).AddDays(1)).ToShortDateString() + " 23:59"
                    $Start = (Get-Date -Year $PlanningYear.Year -Month $i -Day $j)
                    $End = (Get-Date -Year $PlanningYear.Year -Month $i -Day $j)
					$Start = $Start.Date.AddHours(0)
					$End = $End.Date.AddHours(23)
					$End = $End.AddMinutes(59)
					#Write-Host $AppointmentEndTime
                    $Filter = "[MessageClass]='IPM.Appointment' AND [Start] > '"+(Get-Date($Start) -Format g)+"' AND [End] < '"+(Get-Date($End) -Format g)+"'"
                    
                    # Iterate through the filtered appointements to find any existing ones that fit the criteria (proper TimeOfDay tag and not marked as modified on purpose (##)
                    foreach ($Appointment in $Appointments.Restrict($Filter)) {
						
                        if($Appointment.BillingInformation -eq ($SyncManagedFlag + $ToD) -and -not ($Appointment.Subject -like $HashtagRegex)) {  ### To match the exact time: -and ($Appointment.Start -eq [datetime]("$i/$j/" + $PlanningYear.Year + $StartTime) -and ($Appointment.End -eq [datetime]("$i/$j/" + $PlanningYear.Year + $EndTime)))
                            Write-Log -Level INFO -Message ("Existing entry found $i/$j/" + $PlanningYear.Year + " with Subject " + $Appointment.Subject + "  which has been deleted in Excel. Deleting it in Outlook") -logfile $LogFile
                            Write-Host "Deleting " + $Appointment.Subject
                            $Appointment.Delete()
                        }
                    }
                }
            }
        }
    }

    # Close the workbook of this year
    $excel.Workbooks.Close()
}

Write-Log -Level INFO -Message "End Sync $LogStart" -logfile $LogFile


# Remove variables and cleanup, will lead to Excel process running and blocking file if not removed
Remove-Variable -Name Appointment
Remove-Variable -Name Appointments


$excel.Workbooks.Close()
$excel.Quit()

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkSheet)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

Remove-Variable -Name excel

