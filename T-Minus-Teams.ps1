<#
.SYNOPSIS
    Pre-meeting walk-up sequence script for Microsoft Teams.
.DESCRIPTION
    Checks the local Outlook calendar for upcoming Teams meetings. If a meeting is found 
    within the look-ahead window, it checks if the user is marked as Out of Office. 
    If not, it waits until exactly 1 minute before the start time, launches the Teams 
    pre-join lobby, and plays a walk-up track.
#>

# Bulletproof path resolution
$scriptDir = $PSScriptRoot
if (-not $scriptDir) { 
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition 
}
if (-not $scriptDir) { 
    $scriptDir = "C:\Scripts" # Ultimate failsafe
}

$audioPath = Join-Path -Path $scriptDir -ChildPath "theme.wav"
$logPath = Join-Path -Path $scriptDir -ChildPath "T-Minus-Teams.log"
$lockPath = Join-Path -Path $scriptDir -ChildPath "T-Minus-Teams.lock"
$statePath = Join-Path -Path $scriptDir -ChildPath "T-Minus-Teams.state.json"

# Ensure the directory actually exists (fixes the DirectoryNotFoundException)
if (-not (Test-Path $scriptDir)) {
    New-Item -ItemType Directory -Path $scriptDir | Out-Null
}

# Simple logging function
Function Write-Log {
    Param ([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    Add-Content -Path $logPath -Value $logEntry
    Write-Host $logEntry
}

Function Get-TriggerState {
    if (-not (Test-Path $statePath)) { return $null }
    try {
        Get-Content -LiteralPath $statePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    } catch {
        return $null
    }
}

Function Test-AlreadyTriggered {
    Param (
        [string]$EntryId,
        [datetime]$MeetingStart
    )
    if (-not $EntryId) { return $false }
    $s = Get-TriggerState
    if (-not $s) { return $false }
    $startKey = $MeetingStart.ToUniversalTime().ToString("o", [System.Globalization.CultureInfo]::InvariantCulture)
    return ($s.EntryId -eq $EntryId -and $s.MeetingStartUtc -eq $startKey)
}

Function Save-TriggerState {
    Param (
        [string]$EntryId,
        [datetime]$MeetingStart
    )
    $payload = [ordered]@{
        EntryId         = $EntryId
        MeetingStartUtc = $MeetingStart.ToUniversalTime().ToString("o", [System.Globalization.CultureInfo]::InvariantCulture)
        TriggeredAtUtc  = (Get-Date).ToUniversalTime().ToString("o", [System.Globalization.CultureInfo]::InvariantCulture)
    }
    ($payload | ConvertTo-Json -Compress) | Set-Content -LiteralPath $statePath -Encoding UTF8
}

# Windows 10/11 action center toast (no extra modules). Uses the same WinRT API as Settings notifications.
Function Show-TMinusMeetingToast {
    Param (
        [string]$Subject,
        [datetime]$MeetingStart
    )
    try {
        $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
        $null = [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
    } catch {
        Write-Log "Toast unavailable (WinRT types): $($_.Exception.Message)"
        return
    }

    $title = [System.Security.SecurityElement]::Escape('T-Minus Teams')
    $sub = if ($Subject) { [System.Security.SecurityElement]::Escape($Subject) } else { [System.Security.SecurityElement]::Escape('Teams meeting') }
    $when = [System.Security.SecurityElement]::Escape("Starting at $($MeetingStart.ToString('t')) - opening the pre-join lobby.")

    $toastXml = @"
<toast duration="short">
  <visual>
    <binding template="ToastGeneric">
      <text>$title</text>
      <text>$sub</text>
      <text>$when</text>
    </binding>
  </visual>
</toast>
"@

    try {
        $doc = New-Object Windows.Data.Xml.Dom.XmlDocument
        $doc.LoadXml($toastXml.Trim())
        $toast = [Windows.UI.Notifications.ToastNotification]::new($doc)
        # AUMID registered with Windows for desktop PowerShell (matches typical Scheduled Task / hidden runs)
        $appId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show($toast)
        Write-Log "Toast notification shown."
    } catch {
        Write-Log "Toast failed: $($_.Exception.Message)"
    }
}

$lockStream = $null
try {
    try {
        $lockStream = [System.IO.File]::Open(
            $lockPath,
            [System.IO.FileMode]::OpenOrCreate,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )
    } catch {
        Write-Log "Another instance is running (lock in use). Exiting to avoid overlapping runs."
        exit 0
    }

    Write-Log "--- Starting Hype Check ---"

# Hook into Outlook
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder(9)
    $items = $calendar.Items
    
    # CRITICAL: You must sort by Start date BEFORE enabling IncludeRecurrences
    $items.Sort("[Start]")
    $items.IncludeRecurrences = $true
} catch {
    Write-Log "ERROR: Failed to hook into Outlook COM object. Exiting."
    exit
}

# Look ahead 15 minutes to catch the next immediate meeting
$now = Get-Date
$lookAhead = $now.AddMinutes(15)

# Format dates exactly how the Outlook COM object expects them
$strNow = $now.ToString("MM/dd/yyyy hh:mm tt")
$strLookAhead = $lookAhead.ToString("MM/dd/yyyy hh:mm tt")

$filter = "[Start] >= '$strNow' AND [Start] <= '$strLookAhead'"
Write-Log "Scanning calendar with filter: $filter"

$upcomingMeetings = $items.Restrict($filter)

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
    
    # --- Out of Office Check ---
    $isOof = $false
    
    # Format the meeting start time exactly how the COM object expects it
    $strMeetingStart = $meetingStart.ToString("MM/dd/yyyy hh:mm tt")
    
    # Look for any calendar items that overlap with this meeting's start time
    $oofFilter = "[Start] <= '$strMeetingStart' AND [End] >= '$strMeetingStart'"
    Write-Log "Checking for OOO with filter: $oofFilter"
    
    $overlappingItems = $items.Restrict($oofFilter)
    
    foreach ($item in $overlappingItems) {
        # 3 is the internal property value for "Out of Office"
        if ($item.BusyStatus -eq 3) { 
            $isOof = $true
            Write-Log "Detected an Out of Office block ('$($item.Subject)') covering this meeting time."
            break
        }
    }

    if ($isOof) {
        Write-Log "Skipping sequence because you are marked as Out of Office."
    } else {
        # Calculate exactly 1 minute before start time
        $hypeTime = $meetingStart.AddMinutes(-1)
        
        # Wait until the exact hype time
        $timeToWait = New-TimeSpan -Start (Get-Date) -End $hypeTime
        
        if ($timeToWait.TotalSeconds -gt 0) {
            Write-Log "Waiting $($timeToWait.TotalSeconds) seconds until T-minus 1 minute ($hypeTime)..."
            Start-Sleep -Seconds $timeToWait.TotalSeconds
        }

        $entryId = [string]$targetMeeting.EntryID
        if (Test-AlreadyTriggered -EntryId $entryId -MeetingStart $meetingStart) {
            Write-Log "Sequence already ran for this meeting (state file). Skipping duplicate trigger."
        } else {
            Write-Log "Initiating Sequence! Opening Teams lobby and playing audio."

            Show-TMinusMeetingToast -Subject $targetMeeting.Subject -MeetingStart $meetingStart

            Start-Process $teamsUrl
            Save-TriggerState -EntryId $entryId -MeetingStart $meetingStart

            if (Test-Path -LiteralPath $audioPath) {
                $player = New-Object System.Media.SoundPlayer $audioPath
                $player.PlaySync()
            } else {
                Write-Log "WARNING: Audio file not found at $audioPath"
            }
        }
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

} finally {
    if ($null -ne $lockStream) {
        $lockStream.Dispose()
    }
}
