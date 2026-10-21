@echo off
setlocal
rem Use 64-bit Windows PowerShell even when called by 32-bit cmd.exe.
set "PS64=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" (
    set "PS64=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
)
rem Repair the CURRENT USER's DEFAULT Outlook profile; not all test copies.
"%PS64%" -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0Invoke-RDS-OutlookRepair.ps1" -Apply
exit /b %ERRORLEVEL%
