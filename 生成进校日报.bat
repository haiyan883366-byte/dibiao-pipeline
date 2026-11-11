@echo off
setlocal
rem ============================================================
rem  Dibiao Pipeline - one-click daily report generator
rem  Usage:
rem    1) Drag a data source .xlsx onto this file, OR
rem    2) Double-click and paste the source path when prompted
rem  Output: Desktop\DailyReport_YYYYMMDD.xlsx
rem  Works on any Windows PC after running Install.bat once.
rem ============================================================
chcp 65001 >nul

set "SCRIPT=%USERPROFILE%\.workbuddy\skills\dibiao-pipeline\scripts\dibiao_pipeline_v2.py"
set OUTDIR=%USERPROFILE%\Desktop

if "%~1"=="" (
    set /p SRC=Paste source xlsx path, then press Enter:
) else (
    set "SRC=%~1"
)

if not defined SRC goto :empty
if "%SRC%"=="" goto :empty
if not exist "%SRC%" goto :notfound
if not exist "%SCRIPT%" goto :noinstall

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set TODAY=%%i
set "OUT=%OUTDIR%\DailyReport_%TODAY%.xlsx"

echo.
echo Source: %SRC%
echo Output: %OUT%
echo Running... please wait.
echo.

python "%SCRIPT%" "%SRC%" "%OUT%"
if errorlevel 1 goto :fail

echo.
echo ===== DONE =====
echo Output file: %OUT%
start "" explorer.exe /select,"%OUT%"
pause
exit /b 0

:empty
echo [ERROR] No source file given.
pause
exit /b 1

:notfound
echo [ERROR] File not found: %SRC%
pause
exit /b 1

:noinstall
echo [ERROR] Skill not installed yet. Please run Install.bat first.
pause
exit /b 1

:fail
echo [ERROR] Script failed. Check messages above.
pause
exit /b 1
