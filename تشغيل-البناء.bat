@echo off
rem ============================================
rem  Double-click this file to build the
rem  Access correspondence database (.accdb)
rem ============================================
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build-Accdb.ps1"
echo.
pause
