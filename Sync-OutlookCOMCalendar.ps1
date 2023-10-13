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

.EXAMPLE
.\Sync-OutlookCOMCalendar.ps1 -beforeDays 1 -afterDays 7 -destinationAccountEmail "destination@domain.com" -WhatIf
#>

param (
    [int]$beforeDays = 1,
    [int]$afterDays = 7,
    [string]$destinationAccountEmail = "",
    [switch]$WhatIf = $false
)

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

# WhatIf flag
# $WhatIf = $true  # Set to $true to only print data

# Optional destination account parameter
#$destinationAccountEmail = "destination@domain.com"  # Replace with actual email
$destinationCalendar = $null

foreach ($Store in $Namespace.Stores) {
    if ($Store.DisplayName -eq $destinationAccountEmail) {
        $destinationRoot = $Store.GetRootFolder()
        $destinationCalendar = $destinationRoot.Folders.Item("Calendar").Items
        break
    }
}

if ($destinationCalendar -eq $null) {
    Write-Host "Error: Destination calendar could not be found."
    exit 1
}

foreach ($Store in $Namespace.Stores) {
    # Skip the destination account
    if ($Store.DisplayName -eq $destinationAccountEmail) {
        continue
    }

    $rootFolder = $Store.GetRootFolder()
    
    try {
        $calendarFolder = $rootFolder.Folders.Item("Calendar")
    } catch {
        continue
    }

    if ($calendarFolder -ne $null) {
        $calendarItems = $calendarFolder.Items
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
