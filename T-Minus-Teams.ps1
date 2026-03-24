<#
.SYNOPSIS
    Pre-meeting hype sequence script for Microsoft Teams.
.DESCRIPTION
    Checks the local Outlook calendar for upcoming Teams meetings. If a meeting is found 
    within the look-ahead window, it waits until exactly 1 minute before the start time, 
    launches the Teams pre-join lobby, and plays a walk-up track.
#>

$audioPath = "C:\Scripts\T-Minus-Teams\theme.wav"
$logPath = "C:\Scripts\T-Minus-Teams\T-Minus-Teams.log"

# Simple logging function
Function Write-Log {
    Param ([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    Add-Content -Path $logPath -Value $logEntry
    Write-Host $logEntry
}

Write-Log "--- Starting Hype Check ---"

# Hook into Outlook
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder(9)
} catch {
    Write-Log "ERROR: Failed to hook into Outlook COM object. Exiting."
    exit
}

# Look ahead 15 minutes to catch the next immediate meeting
$now = Get-Date
$lookAhead = $now.AddMinutes(15)

$filter = "[Start] >= '$($now.ToString('g'))' AND [Start] <= '$($lookAhead.ToString('g'))'"
$upcomingMeetings = $calendar.Items.Restrict($filter)
$upcomingMeetings.Sort("[Start]")

$targetMeeting = $null
$teamsUrl = $null

# Find the next Teams meeting and extract the join link
foreach ($meeting in $upcomingMeetings) {
    if ($meeting.Body -match "(https://teams\.microsoft\.com/l/meetup-join/[^\s`"<>]+)") {
        $targetMeeting = $meeting
        $teamsUrl = $Matches[1]
        break
    }
}

if ($targetMeeting) {
    $meetingStart = [datetime]$targetMeeting.Start
    Write-Log "Found Teams meeting: '$($targetMeeting.Subject)' starting at $meetingStart"
    
    # Calculate exactly 1 minute before start time
    $hypeTime = $meetingStart.AddMinutes(-1)
    
    # Wait until the exact hype time
    $timeToWait = New-TimeSpan -Start (Get-Date) -End $hypeTime
    
    if ($timeToWait.TotalSeconds -gt 0) {
        Write-Log "Waiting $($timeToWait.TotalSeconds) seconds until T-minus 1 minute ($hypeTime)..."
        Start-Sleep -Seconds $timeToWait.TotalSeconds
    }

    Write-Log "Initiating Sequence! Opening Teams lobby and playing audio."
    
    # Launch the Teams meeting URL
    Start-Process $teamsUrl

    # Blast the theme
    if (Test-Path $audioPath) {
        $player = New-Object System.Media.SoundPlayer $audioPath
        $player.PlaySync()
    } else {
        Write-Log "WARNING: Audio file not found at $audioPath"
    }
} else {
    Write-Log "No upcoming Teams meetings found in the next 15 minutes."
}

# COM Object Cleanup
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($calendar) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Log "--- Check Complete ---"
