@echo off
setlocal

set "ROOT=%~dp0"
set "VENV_PY=%ROOT%.venv\Scripts\python.exe"

if exist "%VENV_PY%" (
  "%VENV_PY%" "%ROOT%launch_compare_app.py"
) else (
  py "%ROOT%launch_compare_app.py"
)

endlocal
