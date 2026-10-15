# Registers (or updates) the Windows scheduled task that runs the daily summary in WSL.
# Daily at 16:00, then every 10 minutes until 23:55 while something is left to write;
# a run missed because the PC was off is started as soon as possible.
$name = "Claude - Resume quotidien"
$action = New-ScheduledTaskAction -Execute "conhost.exe" `
    -Argument "--headless wsl.exe -d Ubuntu -u sulli -- /usr/bin/python3 /home/sulli/.claude/daily-summary/daily_summary.py"
$trigger = New-ScheduledTaskTrigger -Daily -At 16:00
$trigger.Repetition = (New-ScheduledTaskTrigger -Once -At 16:00 `
    -RepetitionInterval (New-TimeSpan -Minutes 10) -RepetitionDuration (New-TimeSpan -Hours 7 -Minutes 55)).Repetition
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 20) -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName $name -Action $action -Trigger $trigger -Settings $settings `
    -Description "Resume quotidien des sessions Claude vers Suivi taches (16h, reessai toutes les 10 min)" -Force
