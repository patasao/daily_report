@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Create a virtual environment and install requirements.txt first.
  exit /b 1
)
".venv\Scripts\python.exe" "main.py" %*
