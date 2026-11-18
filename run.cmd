@echo off
chcp 936 >nul 2>nul
cd /d "%~dp0"
rem 1) double-click      -> process the document currently open in WPS
rem 2) drag a .docx here -> process that file

if "%~1"=="" goto active
powershell -NoProfile -ExecutionPolicy Bypass -File ".\wps-table-unwrap.ps1" -Path "%~1"
goto end

:active
powershell -NoProfile -ExecutionPolicy Bypass -File ".\wps-table-unwrap.ps1"

:end
echo.
echo ================= done =================
pause
