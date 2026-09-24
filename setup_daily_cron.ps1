# setup_daily_cron.ps1
# Registers a Windows Scheduled Task to run the AMFI Daily Pipeline every night at 23:00 (11:00 PM IST)

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$batPath = Join-Path $scriptDir "run_daily_pipeline.bat"
$taskName = "AMFI_Daily_Dashboard_Update"
$scheduledTime = "23:00"

Write-Host "Registering Windows Scheduled Task '$taskName'..." -ForegroundColor Cyan
Write-Host "Target: $batPath"
Write-Host "Schedule: Daily at $scheduledTime IST"

try {
    $action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$batPath`""
    $trigger = New-ScheduledTaskTrigger -Daily -At $scheduledTime
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null
    Write-Host "[SUCCESS] Task '$taskName' successfully registered!" -ForegroundColor Green
    Write-Host "The dashboard will automatically update with fresh AMFI NAVs every night at $scheduledTime."
} catch {
    Write-Warning "PowerShell cmdlet failed: $_. Falling back to schtasks.exe..."
    schtasks /create /tn $taskName /tr "`"$batPath`"" /sc daily /st $scheduledTime /f
}

# Display task info
Get-ScheduledTask -TaskName $taskName | Format-List TaskName, State, Actions, Triggers
