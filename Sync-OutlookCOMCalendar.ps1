<#
.SYNOPSIS
This script syncs calendar items between different Outlook accounts.

.DESCRIPTION
The script fetches calendar items from source calendars and checks if they exist in the destination calendar.
If they do not, it creates them in the destination calendar.

.PARAMETER beforeDays
The number of days before the current date to include calendar items from. Defaults to 1.

.PARAMETER afterDays
The number of days after the current date to include calendar items from. Defaults to 7.

.PARAMETER destinationAccountEmail
The email of the destination account where the calendar items will be created. This is a required parameter.

.PARAMETER WhatIf
When set, the script will only print the actions it would take, without actually making any changes.

.PARAMETER ExtraFolderNames
An array of additional folder names to look for calendar items.

.PARAMETER ExclusionList
An array of calendar items to exclude from the sync process. Defaults to "Public Folders" and "Shared".

.PARAMETER Abbreviation
When set, the script will abbreviate the subject of the calendar items to the first 2 characters of the source store and 6 characters combined from the first two words of the meeting subject.

.EXAMPLE
.\Sync-OutlookCOMCalendar.ps1 -beforeDays 1 -afterDays 7 -destinationAccountEmail "destination@domain.com" -ExtraFolderNames "Folder1", "Folder2" -ExclusionList @("Public Folders", "Shared") -WhatIf -Debug
#>

param (
    [int]$beforeDays = 1,
    [int]$afterDays = 7,
    [string]$destinationAccountEmail = "",
    [switch]$WhatIf = $false,
    [string[]]$ExtraFolderNames = @(),
    [string[]]$ExclusionList = @("Public Folders", "Shared"),
    [switch]$Debug = $false,
    [switch]$Abbreviation = $false
)

if ($Debug) {
    $DebugPreference = 'Continue'
}


# Function to check if Outlook is running
function Check-OutlookRunning {
    $process = Get-Process | Where-Object { $_.ProcessName -eq "OUTLOOK" }
    return ($null -ne $process)
}


# Function to create Outlook COM Object with retry
function Create-OutlookCOM {
    $retryCount = 0
    $Outlook = $null
    while ($retryCount -lt 5 -and ($null -eq $Outlook)) {
        try {
            $Outlook = New-Object -ComObject Outlook.Application
        }
        catch {
            Start-Sleep -Seconds 5
            $retryCount++
        }
    }
    return $Outlook
}

# Check if Outlook is running and start if not
if (-not (Check-OutlookRunning)) {
    Start-Process "outlook.exe"
    Start-Sleep -Seconds 10  # Wait for Outlook to launch
}

# Create Outlook COM Object
$Outlook = Create-OutlookCOM

# Validate COM Object created successfully
if ($null -eq $Outlook) {
    Write-Host "Error: Unable to create Outlook COM Object."
    exit 1
}

$Namespace = $Outlook.GetNamespace("MAPI")

# Time parameters
$startDate = (Get-Date).AddDays(-$beforeDays)
$endDate = (Get-Date).AddDays($afterDays)
 
$destinationCalendar = $null

foreach ($Store in $Namespace.Stores) {
    if ($Store.DisplayName -eq $destinationAccountEmail) {
        $destinationRoot = $Store.GetRootFolder()
        $destinationCalendar = $destinationRoot.Folders.Item("Calendar").Items
        break
    }
}

if ($null -eq $destinationCalendar) {
    Write-Error "Error: Destination calendar could not be found."
    exit 1
}
Write-Debug ("Destination calendar: " + $destinationCalendar.GetType().FullName)

function NextOccurrence($item, $startDate, $endDate) {
    write-debug ("Checking item: " + $item.Subject + ", Start: " + $item.Start + ", End: " + $item.End + ", parameters startDate: " + $startDate + ", endDate: " + $endDate)
    if ($item.IsRecurring) {
        write-debug ("Checking for recurring items: " + $item.Subject)
        $pattern = $item.GetRecurrencePattern()
        $startTime = $pattern.StartTime.TimeOfDay
        write-debug ("Pattern start time: " + $startTime)
        $start = ($startDate).Date.Add($startTime)
        $end = ($startDate).Date.AddDays(14)

        while ($start -le $end) {
            try {
                $nextOccurrence = $pattern.GetOccurrence($start)
                Write-Debug ("Next occurrence: Start: " + $nextOccurrence.Start + ", End: " + $nextOccurrence.End)
                break
            }
            catch {
                # Intentionally unused. I can't test for the next occurence without a call 
                # that might throw an exception that I don't care about at all.
                write-debug ("Exception: " + $_.Exception.Message)
            }
            $start = $start.AddDays(1)
        }

        if ($nextOccurrence) {
            write-debug ("Next occurrence: Start: " + $nextOccurrence.Start + ", End: " + $nextOccurrence.End)
            return $nextOccurrence
        }
    } else {
        write-debug ("Next occurrence: Start: " + $item.Start + ", End: " + $item.End) 
        return $item
    }
}


function ProcessFolder($calendarItems, $WhatIf, $destinationCalendar, $Outlook, $ExclusionList, $startDate, $endDate, $Abbreviation = $false) {
    $calendarItems.IncludeRecurrences = $true
    $calendarItems.Sort("[Start]", $true)

    foreach ($exclude in $ExclusionList) {
        if ($calendarItems.Parent.Name.ToLower() -like ("*" + $exclude.ToLower() + "*")) {
            Write-Debug ("Skipping calendar items from excluded folder: " + $calendarItems.Parent.Name + " based on excludsion match of: " + $exclude)
            return
        }
    }

    Write-Debug "Processing calendar items: ($calendarItems) with count of $($calendarItems.Count)"
    $calendarItems = $calendarItems.Restrict("[Start] >= '$($startDate.ToString("g"))' AND [Start] <= '$($endDate.ToString("g"))'")

    $allItems = @()

    foreach ($item in $calendarItems) {
        write-debug ("Processing item: " + $item.Subject + ", Start: " + $item.Start + ", End: " + $item.End)
        if ($item.IsRecurring) {
            write-debug ("Is Recurring Item: " + $item.Subject)
            $recurringItems = ProcessRecurringItem $item $startDate $endDate
            # Check if $recurringItems is an array and concatenate it with $allItems
            if ($recurringItems -is [System.Array]) {
                $allItems += $recurringItems
            } else {
                $allItems += @($recurringItems)
            }
        } else {
            # Ensure $item is not null before adding to $allItems
            if ($null -ne $item) {
                $allItems += $item
            }
        }
    }


    foreach ($item in $allItems) {
        if($Abbreviation -eq $true) {
            if ($Store.DisplayName.Contains('@')) {
                $item.Subject = ($Store.DisplayName.Split('@')[1].Substring(0, 2)).ToUpper() + " : " + (($item.Subject -split '\s')[0..1] -join '').Substring(0, [Math]::Min(6, (($item.Subject -split '\s')[0..1] -join '').Length))
            } else {
                $item.Subject = $Store.DisplayName.Substring(0, 4).ToUpper() + " : " + (($item.Subject -split '\s')[0..1] -join '').Substring(0, [Math]::Min(5, (($item.Subject -split '\s')[0..1] -join '').Length))
            }
        } else { 
            $item.Subject = $Store.DisplayName + " : " + $item.Subject
        }
        write-debug ("Processing item: " + $item.Subject + ", Start: " + $item.Start + ", End: " + $item.End)
        $existingItems = $destinationCalendar.Restrict("[Subject] = '$($item.Subject)'" + " AND [Start] >= '$($item.Start.ToString("g"))'" + " AND [End] <= '$($item.End.ToString("g"))'")
        write-debug ("Existing items: ")
        foreach ($existingItem in $existingItems) {
            write-debug (" - " + $existingItem.Subject + ", Start: " + $existingItem.Start + ", End: " + $existingItem.End)
        }
        if ($existingItems.Count -eq 0) {
            if ($WhatIf) {
                Write-Host "Would create: $($item.Subject), Start: $($item.Start), End: $($item.End)"
            }
            else {
                $newAppointment = $Outlook.CreateItem(1)
                $newAppointment.Subject = $item.Subject
                $newAppointment.Start = $item.Start
                $newAppointment.End = $item.End
                $newAppointment.Save()
                Write-Host "Created: $($item.Subject)"
            }
        }
        else { 
            Write-Debug ("Item already exists: $($item.Subject)" + ", Start: " + $item.Start + ", End: " + $item.End)
        }
    }
}

function ProcessRecurringItem($item, $startDate, $endDate) {
    $processedItems = @()
    $pattern = $item.GetRecurrencePattern()
    $startTime = $pattern.StartTime.TimeOfDay
    $nextDate = ($startDate).Date.Add($startTime)
    $end = ($endDate).Date.Add($startTime)

    while ($nextDate -le $end) {
        write-debug "Checking occurrence for date: $nextDate"
        try {
            $occurrence = $pattern.GetOccurrence($nextDate)
            if ($occurrence) {
                write-debug ("Processing occurrence: " + $occurrence.Subject + ", Start: " + $occurrence.Start + ", End: " + $occurrence.End)
                $processedItems += $occurrence
            }
        } catch {
            # Exception might be thrown if no occurrence on this date
        }
        $nextDate = $nextDate.AddDays(1)
    }

    # Process exceptions
    foreach ($exception in $pattern.Exceptions) {
        if (!$exception.Deleted) {
            $exceptionItem = $exception.AppointmentItem
            if ($exceptionItem.Start -ge $startDate -and $exceptionItem.Start -le $endDate) {
                write-debug ("Processing exception: " + $exceptionItem.Subject + ", Start: " + $exceptionItem.Start + ", End: " + $exceptionItem.End)
                $processedItems += $exceptionItem
            }
        }
    }

    return $processedItems
}


function RecurseFolders($folder) {
    Write-Debug ("Processing folder: " + $folder.Name)
    
    foreach ($exclude in $ExclusionList) {
        Write-Debug ("Checking exclusion: " + $exclude)
        if ($folder.Name.ToLower() -like ("*" + $exclude.ToLower() + "*")) {
            Write-Debug ("Skipping excluded folder: " + $folder.Name)
            return
        }
    }

    if ($folder.Name -eq "Calendar" -or $folder.Name -in $ExtraFolderNames) {
        Write-Debug ("Processing target folder: " + $folder.Name)
        ProcessFolder $folder.Items $WhatIf $destinationCalendar $Outlook $ExclusionList $startDate $endDate $Abbreviation
    }

    foreach ($subfolder in $folder.Folders) {
        Write-Debug ("Processing subfolder: " + $subfolder.Name)
        RecurseFolders $subfolder
    }
}


foreach ($Store in $Namespace.Stores) {
    Write-Debug ("Processing store: " + $Store.DisplayName)
    $rootFolder = $Store.GetRootFolder()
    RecurseFolders $rootFolder
}