<#
.SYNOPSIS
    Automated installer for the T-Minus-Teams scheduled task.
#>

# 1. Check for Admin rights, and elevate if necessary
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Elevating to Administrator to create the Scheduled Task..."
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$taskName = "T-Minus-Teams"

# 2. Dynamically get the directory this script is running from
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$scriptPath = Join-Path -Path $scriptDir -ChildPath "T-Minus-Teams.ps1"
$audioPath = Join-Path -Path $scriptDir -ChildPath "theme.wav"

Write-Host "Installing $taskName Scheduled Task..." -ForegroundColor Cyan

# 3. Verify the main script is actually in this folder
if (-not (Test-Path $scriptPath)) {
    Write-Error "Could not find T-Minus-Teams.ps1 in $scriptDir. Ensure both scripts are in the same folder."
    Pause
    exit
}

# Warn if the audio file is missing, but don't fail the install
if (-not (Test-Path $audioPath)) {
    Write-Warning "theme.wav not found in $scriptDir. Don't forget to drop your walk-up track in here before your next meeting!"
}

# 4. Remove existing task if it exists (allows for clean upgrades/reinstalls)
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Write-Host "Removing existing $taskName task..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

# 5. Build the Task parameters
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`""
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 10) -RepetitionDuration (New-TimeSpan -Days 3650)
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

# 6. Register the Task
try {
    Register-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -TaskName $taskName -Description "T-Minus-Teams: Plays a walk-up track 1 minute before Teams meetings." -User $env:USERNAME | Out-Null
    Write-Host "Success! The T-Minus-Teams scheduled task has been registered and will check your calendar every 10 minutes." -ForegroundColor Green
} catch {
    Write-Error "Failed to create scheduled task. $_"
}

Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
