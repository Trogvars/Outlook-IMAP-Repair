#requires -Version 5.1
<#
.SYNOPSIS
  Per-user launcher for Repair-OutlookIMAP-v2.ps1 on RDS.
.DESCRIPTION
  Runs in the affected user's session, not as SYSTEM.
  Defaults to the user's default classic Outlook profile.
  -AllProfiles explicitly includes other profiles (including test copies).
  -Apply enables modification; without it, only scans.
  Skips write operations when Outlook or MFCMAPI is open in this session.
  For a multi-session RDS account, ensure Outlook is closed in all sessions.
.NOTES
  Pilot script. Validate on a small group before production rollout.
#>
[CmdletBinding()]
param(
    [switch]$Apply,
    [switch]$AllProfiles,
    [string]$RepairScript = (Join-Path $PSScriptRoot 'Repair-OutlookIMAP-v2.ps1')
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Self-bootstrap when invoked directly from pwsh or 32-bit Windows PowerShell.
if ([IntPtr]::Size -ne 8 -or $PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5) {
    if (-not [Environment]::Is64BitOperatingSystem) {
        throw 'RDS Outlook repair requires 64-bit Windows.'
    }
    $subpath = if ([Environment]::Is64BitProcess) {
        'System32\WindowsPowerShell\v1.0\powershell.exe'
    } else {
        'Sysnative\WindowsPowerShell\v1.0\powershell.exe'
    }
    $ps64 = Join-Path $env:SystemRoot $subpath
    if (-not (Test-Path -LiteralPath $ps64 -PathType Leaf)) {
        throw "Windows PowerShell 5.1 x64 was not found: $ps64"
    }
    $forward = @('-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
    if ($Apply) { $forward += '-Apply' }
    if ($AllProfiles) { $forward += '-AllProfiles' }
    if ($PSBoundParameters.ContainsKey('RepairScript')) { $forward += @('-RepairScript', $RepairScript) }
    & $ps64 @forward
    if ($null -eq $LASTEXITCODE) { exit 1 }
    exit $LASTEXITCODE
}
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
if ($identity.User.Value -eq 'S-1-5-18') {
    throw 'SYSTEM context is unsupported: run in each affected user session.'
}

$logDir = Join-Path $env:LOCALAPPDATA 'OutlookImapRepair'
New-Item -Path $logDir -ItemType Directory -Force | Out-Null
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
$logPath = Join-Path $logDir ("rds-launch-$stamp.log")

function Write-RunLog([string]$Message) {
    $line = ('[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),
        $identity.Name, $Message)
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
    Write-Host $line
}

# FileShare.None prevents simultaneous launches with the same local user profile.
$lock = $null
try {
    try {
        $lock = [System.IO.File]::Open((Join-Path $logDir 'rds-launch.lock'),
            [System.IO.FileMode]::OpenOrCreate,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None)
    }
    catch [System.IO.IOException] {
        Write-RunLog 'DEFER: another repair process already holds the user lock.'
        exit 10
    }

    Write-RunLog ("START Apply={0} AllProfiles={1}" -f $Apply, $AllProfiles)

    if (-not (Test-Path -LiteralPath $RepairScript -PathType Leaf)) {
        Write-RunLog "ERROR: missing core script '$RepairScript'."
        exit 2
    }

    if ($Apply) {
        $currentSession = (Get-Process -Id $PID).SessionId
        $busy = @(Get-Process -Name OUTLOOK,MFCMAPI -ErrorAction SilentlyContinue |
            Where-Object { $_.SessionId -eq $currentSession })
        if ($busy.Count -gt 0) {
            Write-RunLog 'DEFER: Outlook or MFCMAPI is running. Close them and retry or run at next logon.'
            exit 10
        }
    }

    $profileKey = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles'
    $outlookKey = 'HKCU:\Software\Microsoft\Office\16.0\Outlook'
    if (-not (Test-Path -LiteralPath $profileKey)) {
        Write-RunLog 'SKIP: no Outlook Profiles registry key.'
        exit 0
    }

    $available = @(Get-ChildItem -LiteralPath $profileKey |
        ForEach-Object { $_.PSChildName } | Sort-Object -Unique)
    if ($available.Count -eq 0) {
        Write-RunLog 'SKIP: no Outlook profiles.'
        exit 0
    }

    if ($AllProfiles) {
        $selected = $available
    }
    else {
        $defaultProperty = Get-ItemProperty -LiteralPath $outlookKey -Name DefaultProfile -ErrorAction SilentlyContinue
        $defaultProfile = $null
        if ($null -ne $defaultProperty) { $defaultProfile = $defaultProperty.DefaultProfile }
        if ([string]::IsNullOrWhiteSpace($defaultProfile)) {
            if ($available.Count -eq 1) {
                $defaultProfile = $available[0]
            } else {
                Write-RunLog 'ERROR: multiple profiles, but DefaultProfile is missing. Select manually or use -AllProfiles.'
                exit 3
            }
        }
        if ($available -cnotcontains $defaultProfile) {
            Write-RunLog "ERROR: default profile '$defaultProfile' not found."
            exit 3
        }
        $selected = @($defaultProfile)
    }

    $ps64 = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $ps64 -PathType Leaf)) {
        throw "Cannot find x64 Windows PowerShell: $ps64"
    }

    $errors = 0
    foreach ($profile in $selected) {
        Write-RunLog "BEGIN profile '$profile'"
        $argsForCore = @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass',
            '-File',$RepairScript,'-Profile',$profile)
        if ($Apply) { $argsForCore += @('-Apply','-All') }
        $capture = Join-Path $logDir ("rds-core-$stamp-" + [guid]::NewGuid().ToString('N') + '.log')
        try {
            & $ps64 @argsForCore *> $capture
            $code = $LASTEXITCODE
            if ($null -eq $code) { $code = 1 }
            Write-RunLog "END profile '$profile' exit=$code log=$capture"
            if ($code -ne 0) { $errors++ }
        }
        catch {
            $errors++
            Write-RunLog "ERROR profile '$profile': $($_.Exception.Message)"
        }
    }
    Write-RunLog "FINISH profiles=$($selected.Count) errors=$errors."
    if ($errors -gt 0) { exit 1 }
    exit 0
}
finally {
    if ($null -ne $lock) { $lock.Dispose() }
}
