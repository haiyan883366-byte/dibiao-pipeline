@echo off
setlocal
rem ============================================================
rem  Dibiao Pipeline - Installer
rem  Run this ONCE on a new computer. It will:
rem    1) Check Python is available (give install hint if not)
rem    2) Install openpyxl
rem    3) Copy skill files to %USERPROFILE%\.workbuddy\skills\dibiao-pipeline
rem    4) Create a desktop shortcut bat "DailyReport.bat"
rem ============================================================
chcp 65001 >nul
set "SRC=%~dp0"

echo ===== Dibiao Pipeline Installer =====
echo.

rem ---- 1. Check Python ----
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is NOT installed on this PC.
    echo.
    echo Please do these first, then run Install.bat again:
    echo   1. Open https://www.python.org/downloads/ in browser
    echo   2. Download Python 3.x, run installer
    echo   3. IMPORTANT: check the box "Add Python to PATH"
    echo   4. After install, re-run this Install.bat
    echo.
    pause
    exit /b 1
)
for /f %%i in ('python --version') do set PYVER=%%i
echo [OK] Python found: %PYVER%

rem ---- 2. Install openpyxl ----
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Installing openpyxl...
    python -m pip install openpyxl -q
    if errorlevel 1 (
        echo [ERROR] openpyxl install failed. Check network / pip.
        pause
        exit /b 1
    )
)
echo [OK] openpyxl ready

rem ---- 3. Copy skill files ----
set "DEST=%USERPROFILE%\.workbuddy\skills\dibiao-pipeline"
if not exist "%DEST%\scripts" mkdir "%DEST%\scripts"
copy /y "%SRC%scripts\dibiao_pipeline_v2.py" "%DEST%\scripts\" >nul
copy /y "%SRC%scripts\dibiao_pipeline.py" "%DEST%\scripts\" >nul
copy /y "%SRC%SKILL.md" "%DEST%\" >nul
echo [OK] Skill files copied to %DEST%

rem ---- 4. Desktop shortcut bat ----
copy /y "%SRC%生成进校日报.bat" "%USERPROFILE%\Desktop\DailyReport.bat" >nul
echo [OK] Desktop shortcut created: DailyReport.bat

echo.
echo ===== INSTALL DONE =====
echo Usage: drag a source xlsx onto Desktop\DailyReport.bat
echo.

pause
exit /b 0
