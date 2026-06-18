@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\install-sendto.ps1" %*
exit /b %ERRORLEVEL%
