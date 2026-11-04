# Registers the GrantsSeekerAgro scheduled tasks for the current user.
#   Daily search : every day 07:00
#   Weekly email : every Monday 09:00 (after the daily search has finished)
# Tasks run only while you are logged on (required for Outlook and Claude authentication).
# If the PC was off/asleep at the scheduled time, the task runs as soon as possible afterwards.
#
# Run:   powershell -ExecutionPolicy Bypass -File install_tasks.ps1
# Remove: Unregister-ScheduledTask -TaskName "GrantsSeekerAgro*" -Confirm:$false

param(
    [string]$DailyTime  = "07:00",
    [string]$WeeklyDay  = "Monday",
    [string]$WeeklyTime = "09:00"
)

$dir     = Split-Path -Parent $MyInvocation.MyCommand.Path
$script  = Join-Path $dir "grants_seeker.py"
$pyw     = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source
if (-not $pyw) { $pyw = "C:\Python314\pythonw.exe" }

# First-time setup (email, sync folder, topics) if this machine has no config yet
if (-not (Test-Path (Join-Path $dir "config.json"))) {
    Write-Host "No config.json found - running setup first." -ForegroundColor Yellow
    & python $script setup
    if (-not (Test-Path (Join-Path $dir "config.json"))) { Write-Error "Setup was not completed."; exit 1 }
}

$user      = "$env:USERDOMAIN\$env:USERNAME"
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries `
               -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
               -MultipleInstances IgnoreNew

$daily = New-ScheduledTaskAction -Execute $pyw -Argument "`"$script`" daily" -WorkingDirectory $dir
Register-ScheduledTask -TaskName "GrantsSeekerAgro - Daily search" -Force `
    -Description "GrantsSeekerAgro: explores national and international grant calls and updates the Convocatorias OneDrive folder." `
    -Action $daily -Trigger (New-ScheduledTaskTrigger -Daily -At $DailyTime) `
    -Principal $principal -Settings $settings | Out-Null

$weekly = New-ScheduledTaskAction -Execute $pyw -Argument "`"$script`" weekly" -WorkingDirectory $dir
Register-ScheduledTask -TaskName "GrantsSeekerAgro - Weekly email" -Force `
    -Description "GrantsSeekerAgro: sends the weekly grant-calls digest through Outlook." `
    -Action $weekly -Trigger (New-ScheduledTaskTrigger -Weekly -DaysOfWeek $WeeklyDay -At $WeeklyTime) `
    -Principal $principal -Settings $settings | Out-Null

Get-ScheduledTask -TaskName "GrantsSeekerAgro*" |
    Select-Object TaskName, State, @{n="NextRun";e={(Get-ScheduledTaskInfo $_).NextRunTime}} |
    Format-Table -AutoSize
