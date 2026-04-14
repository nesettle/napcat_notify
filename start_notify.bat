@echo off
setlocal

set "ROOT=%~dp0"
set "PYTHON=%ROOT%.venv\Scripts\python.exe"

if not exist "%PYTHON%" (
  set "PYTHON=python"
)

cd /d "%ROOT%"
"%PYTHON%" "%ROOT%notify.py" %*
set "EXIT_CODE=%ERRORLEVEL%"

if "%EXIT_CODE%"=="9009" (
  echo Python was not found. Install Python or create ".venv" first.
)

if not "%EXIT_CODE%"=="0" (
  echo.
  echo notify.py exited with code %EXIT_CODE%.
)

pause
exit /b %EXIT_CODE%
