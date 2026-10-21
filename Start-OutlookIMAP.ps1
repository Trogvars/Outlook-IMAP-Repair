<#
.SYNOPSIS
  Cross-shell bootstrapper for Outlook IMAP repair.
.DESCRIPTION
  Accepts the same parameters as the core repair script and always starts
  Windows PowerShell 5.1 x64, even when invoked from pwsh or 32-bit PowerShell.
  Without -Apply the core performs a read-only scan.
.NOTES
  Use in the affected Windows user's context, not SYSTEM or another account.
  Keep Start-OutlookIMAP.ps1 and Repair-OutlookIMAP-v2.ps1 together.
#>
[CmdletBinding()]
param(
    [switch]$Apply,
    [switch]$All,
    [string]$Profile,
    [ValidatePattern('^[a-fA-F0-9]{32}$')][string]$Uid,
    [string]$BackupRoot
)

$ErrorActionPreference = 'Stop'
if (-not [Environment]::Is64BitOperatingSystem) {
    throw 'Outlook LTSC x64 repair requires 64-bit Windows.'
}
$root = $env:SystemRoot
if ([string]::IsNullOrWhiteSpace($root)) {
    throw 'SystemRoot is not set.'
}
$subpath = if ([Environment]::Is64BitProcess) {
    'System32\WindowsPowerShell\v1.0\powershell.exe'
} else {
    'Sysnative\WindowsPowerShell\v1.0\powershell.exe'
}
$ps64 = Join-Path $root $subpath
$core = Join-Path $PSScriptRoot 'Repair-OutlookIMAP-v2.ps1'
if (-not [System.IO.File]::Exists($ps64)) { throw "Missing x64 Windows PowerShell: $ps64" }
if (-not [System.IO.File]::Exists($core)) { throw "Missing core script: $core" }

$pass = @('-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',$core)
if ($Apply.IsPresent) { $pass += '-Apply' }
if ($All.IsPresent) { $pass += '-All' }
if ($PSBoundParameters.ContainsKey('Profile')) { $pass += @('-Profile',$Profile) }
if ($PSBoundParameters.ContainsKey('Uid')) { $pass += @('-Uid',$Uid) }
if ($PSBoundParameters.ContainsKey('BackupRoot')) { $pass += @('-BackupRoot',$BackupRoot) }

& $ps64 @pass
if ($null -eq $LASTEXITCODE) { exit 1 }
exit $LASTEXITCODE
