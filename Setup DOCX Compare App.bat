@echo off
setlocal

set "ROOT=%~dp0"

where py >nul 2>&1
if %ERRORLEVEL% EQU 0 (
  py -3 "%ROOT%setup_windows.py"
  goto :done
)

where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
  python "%ROOT%setup_windows.py"
  goto :done
)

echo ERROR: Python was not found.
echo Install Python 3.11 or newer, then run this setup again.
exit /b 1

:done
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" (
  echo.
  echo Setup failed. Check "%ROOT%setup_windows.log"
)
exit /b %EXIT_CODE%
