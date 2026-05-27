@echo off
setlocal
cd /d "%~dp0"
set "APP_VENV=%~dp0.venv"
if not exist "%APP_VENV%\Scripts\python.exe" (
  echo Creating local virtual environment...
  python -m venv "%APP_VENV%"
  "%APP_VENV%\Scripts\python.exe" -m pip install -r requirements.txt
)
"%APP_VENV%\Scripts\python.exe" app.py
