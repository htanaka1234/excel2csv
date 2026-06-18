@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\uninstall-sendto.ps1" %*
exit /b %ERRORLEVEL%
