@echo off
rem Double-click launcher for LabRoster.
rem
rem First run creates LabRoster-workbook.xlsx and opens it for you to fill in.
rem Every run after that reads it and writes a timestamped rota, then opens it.
rem
rem Set ROSTER_NO_PAUSE=1 to skip the "press any key" at the end (used by tests).

setlocal
cd /d "%~dp0"
title LabRoster

echo ==========================================
echo   LabRoster
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
if not exist "LabRoster-workbook.xlsx" (
  echo   No input workbook found, so creating one for you...
  echo.
  "%PY%" -m labroster template --out "LabRoster-workbook.xlsx"
  if errorlevel 1 goto :end
  echo.
  echo   Created LabRoster-workbook.xlsx and opening it now.
  echo.
  echo   Add your staff and their competencies, set the roster period on
  echo   the Roster Details sheet, add any leave, then save and close it
  echo   and run this file again to build the draft roster.
  start "" "LabRoster-workbook.xlsx"
  goto :end
)

rem ---- generate, into a timestamped file so nothing is overwritten ----
rem Seconds are included so two runs in the same minute cannot collide.
for /f %%T in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HHmmss"') do set "STAMP=%%T"
set "OUT=LabRoster-draft_%STAMP%.xlsx"

echo   Reading LabRoster-workbook.xlsx
echo   Writing %OUT%
echo.
"%PY%" -m labroster generate --input "LabRoster-workbook.xlsx" --out "%OUT%"
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
