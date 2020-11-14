param (
  [switch] $remove
)

# Read the configuration
. "$PSScriptRoot\config.ps1"

$taskPath = "\letmaik\"
$taskName = "PSExercise"

if ($remove) {
    Unregister-ScheduledTask -TaskPath $taskPath -TaskName $taskName -Confirm:$false
    exit 0
}

# Remove old task if existing
try {
    Get-ScheduledTask -TaskPath $taskPath -TaskName $taskName -ErrorAction Stop | Out-Null
    Unregister-ScheduledTask -TaskPath $taskPath -TaskName $taskName -Confirm:$false
} catch {
}

# Add new task
$scriptRoot = Split-Path $MyInvocation.MyCommand.Path
$runScriptPath = "$scriptRoot\run.ps1"

$psArgs = "-NoLogo -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned"
# Use mshta to avoid showing a console window
# Note that mshta runs in GUI instead of console mode and hence the task "finishes" immediately with exit code 0
# See https://stackoverflow.com/a/45473968
$action = New-ScheduledTaskAction -Execute "%SystemRoot%\system32\mshta.exe" `
    -Argument "vbscript:Execute(`"CreateObject(`"`"Wscript.Shell`"`").Run `"`"powershell ${psArgs} -File `"`"`"`"${runScriptPath}`"`"`"`" `"`", 0 : window.close`")"

$settings = New-ScheduledTaskSettingsSet `
    -Priority 4 `
    -AllowStartIfOnBatteries `
    -DisallowStartOnRemoteAppSession
$trigger = New-ScheduledTaskTrigger -At $startTime -DaysOfWeek $weekDays -Weekly -WeeksInterval 1
$task = Register-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -TaskPath $taskPath -TaskName $taskName
# New-ScheduledTaskTrigger does not offer setting repetition options directly
$task.Triggers.Repetition.Duration = ((Get-Date $stopTime) - (Get-Date $startTime)).ToString("'PT'%h'H'%m'M'")
$task.Triggers.Repetition.Interval = "PT" + $interval
$task | Set-ScheduledTask | Out-Null
