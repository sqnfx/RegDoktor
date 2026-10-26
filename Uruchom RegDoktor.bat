@echo off
rem RegDoktor 11 - uruchamianie (poprosi o uprawnienia administratora)
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0RegDoktor.ps1"
if errorlevel 1 pause
