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

# 2. Get the actual interactive user (Domain\User)
$activeUser = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName
if (-not $activeUser) {
    Write-Error "Could not determine the interactive logged-in user. Exiting."
    Pause
    exit
}

# 3. Dynamically get the directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$scriptPath = Join-Path -Path $scriptDir -ChildPath "T-Minus-Teams.ps1"
$vbsPath    = Join-Path -Path $scriptDir -ChildPath "launcher.vbs"
$audioPath  = Join-Path -Path $scriptDir -ChildPath "theme.wav"

Write-Host "Installing $taskName Scheduled Task for user: $activeUser" -ForegroundColor Cyan

# 4. Verify the required files exist
if (-not (Test-Path $scriptPath) -or -not (Test-Path $vbsPath)) {
    Write-Error "Missing T-Minus-Teams.ps1 or launcher.vbs in $scriptDir. Both are required."
    Pause
    exit
}

if (-not (Test-Path $audioPath)) {
    Write-Warning "theme.wav not found in $scriptDir. Don't forget to drop your walk-up track in here!"
}

# 5. Remove existing task
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Write-Host "Removing existing $taskName task..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

# 6. Build the Task parameters (Now pointing to WScript)
$action = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument "`"$vbsPath`"" -WorkingDirectory $scriptDir
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 10) -RepetitionDuration (New-TimeSpan -Days 3650)
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

# 7. Register the Task
try {
    Register-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -TaskName $taskName -Description "T-Minus-Teams: Plays a walk-up track 1 minute before Teams meetings." -User $activeUser | Out-Null
    Write-Host "Success! The task has been registered and will now run completely invisibly." -ForegroundColor Green
} catch {
    Write-Error "Failed to create scheduled task. $_"
}

Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
