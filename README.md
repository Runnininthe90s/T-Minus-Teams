# T-Minus-Teams

A PowerShell script that runs in the background, automatically finds your next Microsoft Teams meeting, launches the pre-join lobby, and plays your walk-up track exactly one minute before the meeting starts.

## How It Works
1. Hooks into the local Outlook desktop client via COM objects.
2. Scans the calendar for any meetings starting in the next 15 minutes.
3. Uses RegEx to extract the Teams `meetup-join` URL.
4. Pauses execution until exactly **T-minus 60 seconds**.
5. Launches the Teams URL (opening the camera/audio check lobby).
6. Plays a local `.wav` file synchronously.

## Prerequisites
* Windows OS with PowerShell.
* Outlook Desktop Client (must be authenticated and running/syncing).
* A 1-minute `.wav` audio file.

## Setup & Configuration

### 1. File Placement
1. Clone or download this repository to a permanent location (e.g., `C:\Scripts\T-Minus-Teams\`).
2. Rename your 1-minute `.wav` file to `theme.wav` and place it in the same directory.
3. Open `T-Minus-Teams.ps1` and verify the `$audioPath` and `$logPath` variables match your folder structure.

### 2. Automated Task Scheduler Setup
Instead of manually creating the scheduled task through the GUI, open an elevated PowerShell prompt and run the following snippet. 

**Note:** Update the `$scriptPath` variable below if you saved the script somewhere other than `C:\Scripts\T-Minus-Teams\`.

```powershell
$taskName = "T-Minus-Teams"
$scriptPath = "C:\Scripts\T-Minus-Teams\T-Minus-Teams.ps1"

# Define the action: Run PowerShell hidden and bypass execution policy
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""

# Define the trigger: Run once immediately, then repeat every 10 minutes indefinitely
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes 10) -RepetitionDuration (New-TimeSpan -Days 3650)

# Define settings: Allow running on battery
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

# Register the task to run under the current user's context
Register-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -TaskName $taskName -Description "T-Minus-Teams: Plays a hype track 1 minute before Teams meetings." -User $env:USERNAME
