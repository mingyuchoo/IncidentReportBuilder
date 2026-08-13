@echo off
chcp 65001 >nul
REM ============================================================
REM  Jangae MONTHLY report - Web launcher
REM  Opens a local (localhost-only) web UI in the browser.
REM ============================================================

set "PYEXE=%USERPROFILE%\jangae-venv\Scripts\python.exe"

if not exist "%PYEXE%" (
  echo [ERROR] venv Python not found:
  echo         %PYEXE%
  echo         See README for environment reset steps.
  pause
  exit /b 1
)

"%PYEXE%" "%~dp0launcher_monthly.py"
echo.
echo (launcher stopped)
pause
