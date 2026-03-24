# T-Minus-Teams

A PowerShell script that runs in the background, automatically finds your next Microsoft Teams meeting, launches the pre-join lobby, and plays your walk-up track exactly one minute before the meeting starts.

## Features
* **Zero-Touch Execution:** Runs silently in the background via Windows Task Scheduler.
* **Smart PTO Detection:** Checks your local Outlook calendar for `Out of Office` blocks (BusyStatus = 3) and automatically suspends the hype sequence so your PC doesn't blast music in an empty room while you're on vacation.
* **Dynamic Pathing:** Run it from any folder. The script automatically resolves its own directory to find your audio track and write its logs.
* **Precision Timing:** Calculates exactly T-minus 60 seconds from the meeting start time before triggering.
* **Auto-Launch Lobby:** Extracts the hidden Teams URL from the calendar invite and drops you directly into the camera/audio pre-join screen.

## How It Works
1. Hooks into the local Outlook desktop client via COM objects.
2. Scans the calendar for any meetings starting in the next 15 minutes (with recurrence expansion).
3. Verifies you aren't currently marked as Out of Office.
4. Uses RegEx to extract the Teams `meetup-join` URL.
5. Pauses execution until exactly **T-minus 60 seconds**.
6. Launches the Teams URL.
7. Plays a local `.wav` file synchronously.

## Prerequisites
* Windows OS with PowerShell.
* Outlook Desktop Client (must be authenticated and running/syncing).
* A 1-minute `.wav` audio file.

## Setup & Configuration

### 1. File Placement
1. Clone or download this repository to a permanent location (e.g., `C:\Scripts\T-Minus-Teams\`).
2. Rename your 1-minute `.wav` file to `theme.wav` and place it in the exact same directory as the scripts.

### 2. Automated Setup
This repository includes an installer script that will automatically configure the Windows Task Scheduler job for you. 

1. Navigate to the folder where you saved the files.
2. Right-click **`install.ps1`** and select **Run with PowerShell**. 
3. Accept the UAC prompt (the script requires Administrator privileges to register the Scheduled Task and configure the battery/execution policies).
4. The script will automatically build the task to run every 10 minutes in the background.

*Note: If you ever move the folder to a new location, simply run `install.ps1` again. It will cleanly overwrite the old task with the new file paths.*

### 3. Verification
Check the `T-Minus-Teams.log` file in your folder after 10 minutes to verify the script is successfully hooking into Outlook and scanning for meetings.
