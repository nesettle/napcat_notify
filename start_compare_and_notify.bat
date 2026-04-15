@echo off
setlocal

set "ROOT=%~dp0"
set "PYTHON=%ROOT%.venv\Scripts\python.exe"
set "ENTRIES_URL=https://jinshuju.net/forms/SnQ2YZ/entries"
set "PROFILE_DIR=%ROOT%browser_state\jinshuju_profile_qsqni76m"
set "QUALIFICATION_FILE="

for %%I in (C:\Users\Theta\Downloads\?B342~1.XLS) do set "QUALIFICATION_FILE=%%~fI"

if not exist "%PYTHON%" (
  set "PYTHON=python"
)

if not exist "%QUALIFICATION_FILE%" (
  echo Qualification file not found: "%QUALIFICATION_FILE%"
  pause
  exit /b 1
)

cd /d "%ROOT%"
"%PYTHON%" "%ROOT%compare_and_notify.py" --entries-url "%ENTRIES_URL%" --qualification-file "%QUALIFICATION_FILE%" --profile-dir "%PROFILE_DIR%" --headless %*
set "EXIT_CODE=%ERRORLEVEL%"

if "%EXIT_CODE%"=="9009" (
  echo Python was not found. Install Python or create ".venv" first.
)

if not "%EXIT_CODE%"=="0" (
  echo.
  echo compare_and_notify.py exited with code %EXIT_CODE%.
)

pause
exit /b %EXIT_CODE%
