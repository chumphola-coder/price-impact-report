@echo off
setlocal
cd /d "%~dp0"

echo ==================================================
echo   CCWR This-Last Analysis  -  Windows launcher
echo ==================================================
echo.

REM --- 1) Find a working Python (prefer the "py" launcher, then "python") ---
set "PY_CMD="
where py >nul 2>nul && set "PY_CMD=py -3"
if not defined PY_CMD (
  where python >nul 2>nul && set "PY_CMD=python"
)

REM --- 2) If no Python, try to install it automatically with winget ---
if not defined PY_CMD (
  echo Python 3 was not found on this computer.
  echo Trying to install it automatically with Windows Package Manager (winget)...
  echo.
  where winget >nul 2>nul
  if errorlevel 1 (
    echo Could not auto-install because winget is not available on this Windows version.
    echo.
    echo Please install Python 3 manually:
    echo   1. Open https://www.python.org/downloads/windows/
    echo   2. Download "Windows installer ^(64-bit^)".
    echo   3. Run it, TICK "Add python.exe to PATH", then click Install Now.
    echo   4. When it finishes, double-click Run App.bat again.
    echo.
    pause
    exit /b 1
  )
  winget install --id Python.Python.3.12 -e --source winget --accept-package-agreements --accept-source-agreements
  echo.
  echo ------------------------------------------------------------------
  echo   Python has been installed.
  echo   Please CLOSE this window, then double-click Run App.bat again
  echo   so Windows can see the new Python.
  echo ------------------------------------------------------------------
  echo.
  pause
  exit /b 0
)

echo Using Python: %PY_CMD%
echo.

REM --- 3) Create the local environment on first run ---
if not exist .venv (
  echo Creating local Python environment (first run only)...
  %PY_CMD% -m venv .venv
  if errorlevel 1 (
    echo.
    echo ERROR: Failed to create the local Python environment.
    echo Please confirm Python 3 is installed correctly.
    echo.
    pause
    exit /b 1
  )
)

REM --- 4) Install / update the required packages ---
echo Installing / updating required packages...
.venv\Scripts\python -m pip install --upgrade pip >nul 2>nul
.venv\Scripts\python -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo ERROR: Failed to install required packages.
  echo Please check your internet connection or corporate proxy settings.
  echo.
  pause
  exit /b 1
)

REM --- 5) Start the app ---
echo.
echo Opening the app in your browser...
echo (Keep this window open while you use the app. Close it to stop.)
echo.
.venv\Scripts\streamlit run app.py --browser.gatherUsageStats=false
if errorlevel 1 (
  echo.
  echo ERROR: The app failed to start.
  echo.
  pause
  exit /b 1
)

pause
