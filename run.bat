@echo off
rem Double-click to open the cohort shipping menu.
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 (
  echo Python was not found. Install it from the Microsoft Store or python.org, then try again.
  pause
  exit /b 1
)
py cohort.py %*
if errorlevel 1 pause
