@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
rem גרור על הקובץ הזה את האקסל ו/או צילום המסך / המייל של ברינקס (שמור כ-HTML). מה שלא גררת נמשך מ-Outlook.
if "%~1"=="" (
  echo גרור על הקובץ הזה את תכנית ההתקנות ^(xlsx^) ו/או את צילום המייל של ברינקס ^(png/jpg^) או את המייל שנשמר כ-htm.
  pause
  exit /b
)
set ARGS=
set IMGS=
:loop
if "%~1"=="" goto run
set "EXT=%~x1"
if /i "%EXT%"==".xlsx" set ARGS=%ARGS% --plan "%~1"
if /i "%EXT%"==".xlsm" set ARGS=%ARGS% --plan "%~1"
if /i "%EXT%"==".xls"  set ARGS=%ARGS% --plan "%~1"
if /i "%EXT%"==".htm"  set ARGS=%ARGS% --email-html "%~1"
if /i "%EXT%"==".html" set ARGS=%ARGS% --email-html "%~1"
if /i "%EXT%"==".png"  set IMGS=%IMGS% "%~1"
if /i "%EXT%"==".jpg"  set IMGS=%IMGS% "%~1"
if /i "%EXT%"==".jpeg" set IMGS=%IMGS% "%~1"
shift
goto loop
:run
if defined IMGS set ARGS=%ARGS% --email-image %IMGS%
python sync_schedule.py --force %ARGS%
pause
