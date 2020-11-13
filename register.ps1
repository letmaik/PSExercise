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

$commonPSArgs = "-NoLogo -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned"
if ($retryCount -eq 0) {
    # Use mshta to avoid showing a console window
    # See https://stackoverflow.com/a/45473968
    $action = New-ScheduledTaskAction -Execute "%SystemRoot%\system32\mshta.exe" `
        -Argument "vbscript:Execute(`"CreateObject(`"`"Wscript.Shell`"`").Run `"`"powershell ${commonPSArgs} -File `"`"`"`"${runScriptPath}`"`"`"`" `"`", 0 : window.close`")"
} else {
    # Cannot use mshta to hide the console window as the exit code
    # would get lost and we need it for task retry to work ("No" exits with 1)
    # "-windowstyle hidden" flashes the console window shortly
    $action = New-ScheduledTaskAction -Execute "powershell" -Argument "${commonPSArgs} -WindowStyle hidden -File `"${runScriptPath}`""
}
$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
    -MultipleInstances IgnoreNew `
    -Priority 4 `
    -AllowStartIfOnBatteries `
    -DisallowStartOnRemoteAppSession
$trigger = New-ScheduledTaskTrigger -Daily -At $startTime
$task = Register-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -TaskPath $taskPath -TaskName $taskName
# New-ScheduledTaskTrigger does not offer setting repetition options directly
$task.Triggers.Repetition.Duration = ((Get-Date $stopTime) - (Get-Date $startTime)).ToString("'PT'%h'H'%m'M'")
$task.Triggers.Repetition.Interval = "PT" + $interval
if ($retryCount -gt 0) {
    # New-ScheduledTaskSettingsSet can also be used to set these,
    # but using this method allows to use ISO duration strings directly
    $task.Settings.RestartCount = $retryCount
    $task.Settings.RestartInterval = "PT" + $retryInterval
}
$task | Set-ScheduledTask | Out-Null
