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

.EXAMPLE
.\Sync-OutlookCOMCalendar.ps1 -beforeDays 1 -afterDays 7 -destinationAccountEmail "destination@domain.com" -ExtraFolderNames "Folder1", "Folder2" -ExclusionList @("Public Folders", "Shared") -WhatIf -Debug
#>

param (
    [int]$beforeDays = 1,
    [int]$afterDays = 7,
    [string]$destinationAccountEmail = "",
    [switch]$WhatIf = $false,
    [string[]]$ExtraFolderNames = @(),
    [string[]]$ExclusionList = @("Public Folders","Shared"),
    [switch]$Debug = $false
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
        } catch {
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

function ProcessFolder($calendarItems, $WhatIf, $destinationCalendar, $Outlook, $ExclusionList) {
    $shouldProcess = $true
    foreach ($exclude in $ExclusionList) {
        if ($calendarItems.Parent.Name.ToLower() -like ("*" + $exclude.ToLower() + "*")) {
            Write-Debug ("Skipping calendar items from folder: " + $calendarItems.Parent.Name)
            $shouldProcess = $false
            break
        }
    }
    if ($shouldProcess) {
        Write-Debug "Processing calendar items: ($calendarItems) with count of $($calendarItems.Count)"
        $calendarItems = $calendarItems.Restrict("[Start] >= '$($startDate.ToString("g"))' AND [Start] <= '$($endDate.ToString("g"))'")
        foreach ($item in $calendarItems) {
            $existingItems = $destinationCalendar.Restrict("[Subject] = '$($item.Subject)'")
            if ($existingItems.Count -eq 0) {
                if ($WhatIf) {
                    Write-Host "Would create: $($item.Subject), Start: $($item.Start), End: $($item.End)"
                } else {
                    $newAppointment = $Outlook.CreateItem(1)
                    $newAppointment.Subject = $item.Subject
                    $newAppointment.Start = $item.Start
                    $newAppointment.End = $item.End
                    $newAppointment.Save()
                    Write-Host "Created: $($item.Subject)"
                }
            }
        }
    }
}


function RecurseFolders($folder) {
    Write-Debug ("Processing folder: " + $folder.Name)
    
    $shouldSkip = $false
    foreach ($exclude in $ExclusionList) {
        if ($folder.Name.ToLower() -like ("*" + $exclude.ToLower() + "*")) {
            $shouldSkip = $true
            break
        }
    }
    
    if ($shouldSkip) {
        Write-Debug ("Skipping excluded folder: " + $folder.Name)
        return
    }

    if ($folder.Name -eq "Calendar") {
        foreach ($subfolder in $folder.Folders) {
            Write-Debug ("Found a subfolder to process: " + $subfolder.Name)
            ProcessFolder $subfolder.Items $WhatIf $destinationCalendar $Outlook $ExclusionList # Passing arguments
        }
    } elseif ($folder.Name -in $ExtraFolderNames) {
        Write-Debug ("Found an ExtraFolderName folder to process: " + $folder.Name)
        ProcessFolder $folder.Items $WhatIf $destinationCalendar $Outlook $ExclusionList # Passing arguments
    }

    foreach ($subfolder in $folder.Folders) {
        Write-Debug ("Processing subfolder"+ $subfolder.Name )
        RecurseFolders $subfolder
    }
}


foreach ($Store in $Namespace.Stores) {
    Write-Debug ("Processing store: " + $Store.DisplayName)
    $rootFolder = $Store.GetRootFolder()
    RecurseFolders $rootFolder
}