@echo off
echo EntryPass-MiHCM Sync
echo.
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0sync.ps1"
echo.
echo Sync complete.
pause
