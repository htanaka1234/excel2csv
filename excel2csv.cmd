@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\excel2csv.ps1" %*
exit /b %ERRORLEVEL%
