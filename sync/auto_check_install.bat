@echo off
chcp 65001 >nul
cd /d "%~dp0"
for /f "delims=" %%P in ('where pythonw') do set PYW=%%P& goto found
echo pythonw not found. Install Python from python.org first.& pause & exit /b
:found
schtasks /Create /TN "LZ Install Schedule" /SC DAILY /ST 07:00 /RI 30 /DU 12:00 /TR "\"%PYW%\" \"%~dp0sync_schedule.py\" --auto --no-open" /F
echo.
echo Done: the check runs every 30 minutes between 07:00 and 19:00.
echo To remove: schtasks /Delete /TN "LZ Install Schedule" /F
pause
