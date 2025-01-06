# Define Variables
# Path to the folder containing video files
$videoFolderPath = "C:\Users\UserName\Videos"  # <-- Replace with your actual path

# Define time window (in minutes) to match the video time with calendar event
$timeWindowMinutes = 1

# Create Outlook COM object with error handling
try {
    $outlook = New-Object -ComObject Outlook.Application
} catch {
    Write-Host "Error: Outlook is not installed or accessible. Please ensure Outlook is installed and configured properly." -ForegroundColor Red
    exit
}

# Get MAPI namespace to access Outlook items with error handling
try {
    $namespace = $outlook.GetNamespace("MAPI")
    $calendarFolder = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
} catch {
    Write-Host "Error: Unable to access Outlook calendar. Please ensure that Outlook is properly configured." -ForegroundColor Red
    exit
}

# Get all .mkv files in the specified folder with error handling for invalid or empty directories
try {
    $videoFiles = Get-ChildItem -Path $videoFolderPath -Filter "*.mkv"
    if ($videoFiles.Count -eq 0) {
        Write-Host "Error: No video files found in the specified folder." -ForegroundColor Yellow
        exit
    }
} catch {
    Write-Host "Error: The specified folder path is invalid or inaccessible." -ForegroundColor Red
    exit
}

# Loop through each video file
foreach ($file in $videoFiles) {
    # Use regex to extract the date and time from the file name
    if ($file.Name -match "^(\d{4}-\d{2}-\d{2}) (\d{2})-(\d{2})-(\d{2})\.mkv$") {
        # Capture the date part (YYYY-MM-DD) and time part (HH-MM-SS)
        $dateString = $matches[1]
        $timeString = "$($matches[2]):$($matches[3]):$($matches[4])"

        # Parse the date and time into a DateTime object
        try {
            $fileDateTime = [datetime]::ParseExact("$dateString $timeString", "yyyy-MM-dd HH:mm:ss", $null)
        } catch {
            Write-Host "Error: Unable to parse date and time from file name '$($file.Name)'." -ForegroundColor Red
            continue
        }

        # Define a time window around the file date and time (e.g., +/- 1 minute)
        $startTime = $fileDateTime.AddMinutes(-$timeWindowMinutes)
        $endTime = $fileDateTime.AddMinutes($timeWindowMinutes)

        # Format the times for the Outlook query (MM/dd/yyyy hh:mm tt)
        $startTimeString = $startTime.ToString("MM/dd/yyyy hh:mm tt")
        $endTimeString = $endTime.ToString("MM/dd/yyyy hh:mm tt")

        # Build the restriction query to find calendar items overlapping with the video time
        $restriction = "[Start] <= '$endTimeString' AND [End] >= '$startTimeString'"

        # Retrieve all calendar items and include recurring events, limit the date range to improve performance
        try {
            $items = $calendarFolder.Items
            $items.IncludeRecurrences = $true
            $items.Sort("[Start]")
            $items = $items.Restrict("[Start] >= '$($startTime.AddDays(-1).ToString("MM/dd/yyyy hh:mm tt"))' AND [End] <= '$($endTime.AddDays(1).ToString("MM/dd/yyyy hh:mm tt"))'")
        } catch {
            Write-Host "Error: Unable to retrieve calendar items." -ForegroundColor Red
            continue
        }

        # Apply the restriction to filter matching calendar items
        $matchingItems = $items.Restrict($restriction)

        # Fix potential issue with case sensitivity by converting to lowercase
        if ($matchingItems.Count -gt 0) {
            # Loop through matching items to find the closest match by comparing date and time
            $closestItem = $null
            $smallestDifference = [timespan]::MaxValue
            foreach ($calendarItem in $matchingItems) {
                $itemStart = [datetime]$calendarItem.Start
                $timeDifference = [timespan]::FromTicks(($itemStart - $fileDateTime).Ticks)
                if ($timeDifference -lt $smallestDifference) {
                    $smallestDifference = $timeDifference
                    $closestItem = $calendarItem
                }
            }

            if ($closestItem -ne $null) {
                # Get the subject of the closest calendar event
                $subject = $closestItem.Subject

                # Sanitize the subject to remove any invalid file name characters
                $safeSubject = [RegEx]::Replace($subject, '[<>:"/\\|?*]', '-')

                # Construct the new file name with the appended subject
                $newFileName = "$($file.BaseName) - $safeSubject$($file.Extension)"

                # Perform the file rename
                try {
                    Rename-Item -Path $file.FullName -NewName $newFileName
                    Write-Host "Renamed '$($file.Name)' to '$newFileName'" -ForegroundColor Green
                } catch {
                    Write-Host "Error: Unable to rename file '$($file.Name)'." -ForegroundColor Red
                }
            } else {
                Write-Host "No closest matching calendar event found for '$($file.Name)'" -ForegroundColor Yellow
            }
        } else {
            # No matching calendar event found for the video file
            Write-Host "No calendar event found for '$($file.Name)'" -ForegroundColor Yellow
        }
    } else {
        # File name does not match the expected pattern
        Write-Host "File name '$($file.Name)' does not match the expected pattern." -ForegroundColor Yellow
    }
}

# Output completion message to the console
Write-Host "File rename operation completed."