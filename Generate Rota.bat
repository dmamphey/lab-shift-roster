@echo off
rem Double-click launcher for the lab shift roster generator.
rem
rem First run creates roster_input.xlsx and opens it for you to fill in.
rem Every run after that reads it and writes a timestamped rota, then opens it.
rem
rem Set ROSTER_NO_PAUSE=1 to skip the "press any key" at the end (used by tests).

setlocal
cd /d "%~dp0"
title Lab Shift Roster

echo ==========================================
echo   Lab Shift Roster
echo ==========================================
echo.

rem ---- find a working Python -------------------------------------------
set "PY="
py -c "import sys" >nul 2>&1
if not errorlevel 1 set "PY=py"

if not defined PY (
  python -c "import sys" >nul 2>&1
  if not errorlevel 1 set "PY=python"
)

if not defined PY (
  if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
  )
)

if not defined PY (
  echo   Python was not found on this computer.
  echo.
  echo   Install it from https://www.python.org/downloads/
  echo   During setup, tick "Add python.exe to PATH", then run this again.
  goto :end
)

rem ---- make sure openpyxl is available --------------------------------
"%PY%" -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
  echo   Installing the openpyxl library, one moment...
  "%PY%" -m pip install --quiet --disable-pip-version-check openpyxl
  if errorlevel 1 (
    echo.
    echo   Could not install openpyxl. Check your internet connection,
    echo   then run this again.
    goto :end
  )
  echo   Done.
  echo.
)

rem ---- first run: no input workbook yet -------------------------------
if not exist "roster_input.xlsx" (
  echo   No input workbook found, so creating one for you...
  echo.
  "%PY%" "lab_roster.py" template --out "roster_input.xlsx"
  if errorlevel 1 goto :end
  echo.
  echo   Created roster_input.xlsx and opening it now.
  echo.
  echo   Replace the example staff with your real staff, set the date
  echo   range on the Settings sheet, add any leave, then save and close
  echo   it and run this file again to build the rota.
  start "" "roster_input.xlsx"
  goto :end
)

rem ---- generate, into a timestamped file so nothing is overwritten ----
rem Seconds are included so two runs in the same minute cannot collide.
for /f %%T in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HHmmss"') do set "STAMP=%%T"
set "OUT=rota_%STAMP%.xlsx"

echo   Reading roster_input.xlsx
echo   Writing %OUT%
echo.
"%PY%" "lab_roster.py" generate --input "roster_input.xlsx" --out "%OUT%"
if errorlevel 1 (
  echo.
  echo   Something went wrong - the message above says what.
  goto :end
)

echo.
echo   Opening %OUT%
echo.
echo   Check the Summary sheet: the rule check block near the bottom
echo   tells you if any shift ended up short of cover.
start "" "%OUT%"

:end
echo.
if not defined ROSTER_NO_PAUSE pause
endlocal
