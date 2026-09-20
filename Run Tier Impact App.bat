@echo off
setlocal
cd /d "%~dp0"

echo ==================================================
echo   Service Tier Change Impact  -  Windows launcher
echo ==================================================
echo.

set "PY_CMD="
where py >nul 2>nul && set "PY_CMD=py -3"
if not defined PY_CMD (
  where python >nul 2>nul && set "PY_CMD=python"
)

if not defined PY_CMD (
  echo Python 3 was not found on this computer.
  echo Please install Python 3 first ^(see README.md^), then run this again.
  echo.
  pause
  exit /b 1
)

echo Using Python: %PY_CMD%
echo.

if not exist .venv (
  echo Creating local Python environment (first run only)...
  %PY_CMD% -m venv .venv
  if errorlevel 1 (
    echo.
    echo ERROR: Failed to create the local Python environment.
    echo.
    pause
    exit /b 1
  )
)

echo Installing / updating required packages...
.venv\Scripts\python -m pip install --upgrade pip >nul 2>nul
.venv\Scripts\python -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo ERROR: Failed to install required packages.
  echo.
  pause
  exit /b 1
)

echo.
echo Opening the app in your browser...
echo (Keep this window open while you use the app. Close it to stop.)
echo.
.venv\Scripts\streamlit run tier_impact_app.py --browser.gatherUsageStats=false
if errorlevel 1 (
  echo.
  echo ERROR: The app failed to start.
  echo.
  pause
  exit /b 1
)

pause
