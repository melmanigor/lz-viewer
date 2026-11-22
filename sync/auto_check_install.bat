@echo off
chcp 65001 >nul
cd /d "%~dp0"
for /f "delims=" %%P in ('where pythonw') do set PYW=%%P& goto found
echo pythonw not found. Install Python from python.org first.& pause & exit /b
:found
schtasks /Create /TN "LZ Install Schedule" /SC DAILY /ST 07:00 /RI 30 /DU 12:00 /TR "\"%PYW%\" \"%~dp0sync_schedule.py\" --auto --no-open" /F
rem If the PC was off at a scheduled time, run as soon as it is on again
powershell -NoProfile -Command "$s = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries; Set-ScheduledTask -TaskName 'LZ Install Schedule' -Settings $s | Out-Null"
echo.
echo Done: the check runs every 30 minutes between 07:00 and 19:00,
echo right after the computer is turned on if a check was missed.
echo To remove: schtasks /Delete /TN "LZ Install Schedule" /F
pause
