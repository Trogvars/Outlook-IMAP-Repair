@echo off
setlocal DisableDelayedExpansion
rem Universal launcher for classic Outlook MAPI repair (current Windows user).
rem Works when called from x86/x64 cmd.exe, Windows PowerShell, or pwsh.
rem Default (no arguments): READ-ONLY scan.
rem Examples:
rem   Run-OutlookIMAP.cmd
rem   Run-OutlookIMAP.cmd -Profile "Outlook"
rem   Run-OutlookIMAP.cmd -Apply -Profile "Outlook" -All
rem   Run-OutlookIMAP.cmd -Apply -Profile "Outlook" -Uid "0123456789ABCDEF0123456789ABCDEF"

set "REPAIR=%~dp0Repair-OutlookIMAP-v2.ps1"
set "PS64=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
rem Sysnative only exists for 32-bit processes on 64-bit Windows (WOW64).
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" (
    set "PS64=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
)

if not exist "%REPAIR%" (
    echo ERROR: Core script not found: "%REPAIR%" 1>&2
    endlocal & exit /b 2
)
if not exist "%PS64%" (
    echo ERROR: Windows PowerShell 5.1 executable not found: "%PS64%" 1>&2
    endlocal & exit /b 3
)

"%PS64%" -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%REPAIR%" %*
set "RESULT=%ERRORLEVEL%"
endlocal & exit /b %RESULT%
