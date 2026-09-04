@echo off
setlocal
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0pdf2md.ps1" %*
exit /b %ERRORLEVEL%
