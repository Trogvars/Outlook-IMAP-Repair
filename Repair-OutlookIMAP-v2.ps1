#requires -Version 5.1
<#
.SYNOPSIS
  Outlook LTSC 2021 x64 IMAP (INTERSTOR) profile repair - pilot, not production validated.
.DESCRIPTION
  Without parameters: scans CURRENT USER's classic Outlook profiles (READ-ONLY).
  -Apply -Profile <name> -Uid <32 hex>: repairs precisely one matching INTERSTOR service.
  -Apply -Profile <name> -All: repairs all matching services in ONE specified profile.
  -All first collects targets in a read-only pass; modifications are performed
  separately by UID after a single mandatory registry backup.
  Only when BOTH: PR_SERVICE_ENTRY_NAME_A='PRXServiceEntry@40' AND
  PR_SERVICE_SUPPORT_FILES_A=PT_ERROR(MAPI_E_NOT_FOUND).
  Writes both correct properties in one Extended MAPI SetProps call:
   PR_SERVICE_ENTRY_NAME_A = PRXServiceEntry
   PR_SERVICE_SUPPORT_FILES_A = [pstprx.dll] (PT_MV_STRING8)

  Single-mailbox repair has been user-tested with Outlook LTSC 2021 x64.
  Bulk mode has NOT yet been field-tested; test one profile before broad rollout.
  Never run elevated as another user/SYSTEM. Outlook and MFCMAPI must be closed for -Apply.
#>
[CmdletBinding()]
param(
    [switch]$Apply,
    [switch]$All,
    [string]$Profile,
    [ValidatePattern('^[A-Fa-f0-9]{32}$')][string]$Uid,
    [string]$BackupRoot = (Join-Path $env:LOCALAPPDATA 'OutlookImapRepair')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Self-bootstrap when launched from pwsh or 32-bit Windows PowerShell.
# The Extended MAPI interop below still runs only under Windows PowerShell 5.1 x64.
if ([IntPtr]::Size -ne 8 -or $PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5) {
    if (-not [Environment]::Is64BitOperatingSystem) {
        throw 'Outlook LTSC x64 repair requires 64-bit Windows.'
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
    if ($All) { $forward += '-All' }
    if ($PSBoundParameters.ContainsKey('Profile')) { $forward += @('-Profile', $Profile) }
    if ($PSBoundParameters.ContainsKey('Uid')) { $forward += @('-Uid', $Uid) }
    if ($PSBoundParameters.ContainsKey('BackupRoot')) { $forward += @('-BackupRoot', $BackupRoot) }
    & $ps64 @forward
    if ($null -eq $LASTEXITCODE) { exit 1 }
    exit $LASTEXITCODE
}
if ($Apply -and [Security.Principal.WindowsIdentity]::GetCurrent().User.Value -eq 'S-1-5-18') {
    throw 'Refusing to modify MAPI profiles as SYSTEM. Run as the affected user.'
}
if ($All -and -not $Apply) { throw '-All is only valid together with -Apply.' }
if ($All -and $Uid) { throw '-All and -Uid are mutually exclusive.' }
if ($Apply -and ([string]::IsNullOrWhiteSpace($Profile) -or (-not $All -and [string]::IsNullOrWhiteSpace($Uid)))) {
    throw 'Use either -Apply -Profile "Outlook" -Uid 32HEX or -Apply -Profile "Outlook" -All.'
}
if ($Uid -and [string]::IsNullOrWhiteSpace($Profile)) { throw '-Uid requires -Profile.' }

$cs = @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

public static class OutlookMapiInterop
{
    // All pointers / offsets below are for Win64 Extended MAPI. Keep x86 disabled.
    private const uint PR_DISPLAY_NAME_A = 0x3001001E;
    private const uint PR_SERVICE_NAME_A = 0x3D09001E;
    private const uint PR_SERVICE_UID = 0x3D0C0102;
    private const uint PR_SERVICE_ENTRY_NAME_A = 0x3D0B001E;
    private const uint PR_SERVICE_SUPPORT_FILES_A = 0x3D0F101E;
    private const uint PR_SUPPORT_ERROR = 0x3D0F000A;
    private const uint MAPI_MODIFY = 0x00000001;
    private const uint MAPI_FORCE_ACCESS = 0x00080000;
    private const int E_NOT_FOUND = unchecked((int)0x8004010F);
    private const int PROP_SIZE = 24;
    private const int ROW_SIZE = 16;

    public sealed class Result
    {
        public string Profile { get; set; }
        public string Service { get; set; }
        public string Uid { get; set; }
        public string Entry { get; set; }
        public string Support { get; set; }
        public string Status { get; set; }
        public string Detail { get; set; }
    }

    [DllImport("mapi32.dll", EntryPoint="MAPIInitialize", CallingConvention=CallingConvention.StdCall)]
    private static extern int MAPIInitialize(IntPtr init);
    [DllImport("mapi32.dll", EntryPoint="MAPIUninitialize", CallingConvention=CallingConvention.StdCall)]
    private static extern void MAPIUninitialize();
    [DllImport("mapi32.dll", EntryPoint="MAPIAdminProfiles", CallingConvention=CallingConvention.StdCall)]
    private static extern int MAPIAdminProfiles(uint flags, out IntPtr profAdmin);
    [DllImport("mapi32.dll", EntryPoint="MAPIFreeBuffer", CallingConvention=CallingConvention.StdCall)]
    private static extern int MAPIFreeBuffer(IntPtr ptr);
    [DllImport("mapi32.dll", EntryPoint="FreeProws", CallingConvention=CallingConvention.StdCall)]
    private static extern void FreeProws(IntPtr rows);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TGetTable(IntPtr self, uint flags, out IntPtr table);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TAdminServices(IntPtr self, [MarshalAs(UnmanagedType.LPStr)] string name,
                                       IntPtr password, IntPtr hwnd, uint flags, out IntPtr admin);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TSetColumns(IntPtr self, IntPtr tags, uint flags);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TQueryRows(IntPtr self, int count, uint flags, out IntPtr rows);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TOpenSect(IntPtr self, IntPtr uid, IntPtr iid, uint flags, out IntPtr section);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TGetProps(IntPtr self, IntPtr tags, uint flags, out uint count, out IntPtr values);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TSetProps(IntPtr self, uint count, IntPtr values, out IntPtr problems);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int TSaveChanges(IntPtr self, uint flags);

    private static T Fn<T>(IntPtr obj, int index) where T:class
    {
        if (obj == IntPtr.Zero) throw new InvalidOperationException("Null MAPI COM interface");
        IntPtr vtable = Marshal.ReadIntPtr(obj);
        return (T)(object)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(vtable, index * IntPtr.Size), typeof(T));
    }
    private static void Check(int hr, string stage)
    {
        if (hr < 0) throw new InvalidOperationException(stage + " failed: 0x" + ((uint)hr).ToString("X8"));
    }
    private static void Release(ref IntPtr p)
    {
        if (p != IntPtr.Zero) { Marshal.Release(p); p = IntPtr.Zero; }
    }
    private static string A(IntPtr prop)
    {
        IntPtr p = Marshal.ReadIntPtr(prop, 8);
        return p == IntPtr.Zero ? "" : Marshal.PtrToStringAnsi(p);
    }
    private static uint Tag(IntPtr prop) { return unchecked((uint)Marshal.ReadInt32(prop, 0)); }
    private static IntPtr Prop(IntPtr data, uint count, uint propertyId)
    {
        if (count > 512) throw new InvalidOperationException("Suspiciously large MAPI property count");
        for (int i = 0; i < count; i++) {
            IntPtr p = IntPtr.Add(data, i * PROP_SIZE);
            if ((Tag(p) >> 16) == propertyId) return p;
        }
        return IntPtr.Zero;
    }
    private static IntPtr NewTags(params uint[] tags)
    {
        IntPtr p = Marshal.AllocHGlobal(4 + 4 * tags.Length);
        Marshal.WriteInt32(p, 0, tags.Length);
        for (int i = 0; i < tags.Length; i++) Marshal.WriteInt32(p, 4 + i * 4, unchecked((int)tags[i]));
        return p;
    }
    private static string UidHex(IntPtr binary)
    {
        if (binary == IntPtr.Zero || Tag(binary) != PR_SERVICE_UID) return null;
        if (Marshal.ReadInt32(binary, 8) != 16) return null;
        IntPtr addr = Marshal.ReadIntPtr(binary, 16);
        if (addr == IntPtr.Zero) return null;
        byte[] bytes = new byte[16]; Marshal.Copy(addr, bytes, 0, 16);
        return BitConverter.ToString(bytes).Replace("-", "");
    }
    private static string GetMultiA(IntPtr prop)
    {
        if (prop == IntPtr.Zero || Tag(prop) != PR_SERVICE_SUPPORT_FILES_A) return null;
        int count = Marshal.ReadInt32(prop, 8);
        IntPtr arr = Marshal.ReadIntPtr(prop, 16);
        if (count < 0 || count > 100 || (count > 0 && arr == IntPtr.Zero)) return "<invalid>";
        List<string> values = new List<string>();
        for (int i = 0; i < count; i++) {
            IntPtr q = Marshal.ReadIntPtr(arr, i * IntPtr.Size);
            values.Add(q == IntPtr.Zero ? "<null>" : Marshal.PtrToStringAnsi(q));
        }
        return String.Join(";", values.ToArray());
    }
    private static Result Visit(IntPtr admin, string name, string service, string uid,
                                bool apply)
    {
        Result r = new Result { Profile = name, Service = service, Uid = uid, Status="ERROR" };
        IntPtr section=IntPtr.Zero, uidPtr=IntPtr.Zero, tags=IntPtr.Zero, vals=IntPtr.Zero;
        try {
            byte[] uidBytes = new byte[16];
            for (int i=0; i<16; i++) uidBytes[i] = Convert.ToByte(uid.Substring(i*2, 2), 16);
            uidPtr=Marshal.AllocHGlobal(16); Marshal.Copy(uidBytes,0,uidPtr,16);
            tags=NewTags(PR_SERVICE_ENTRY_NAME_A, PR_SERVICE_SUPPORT_FILES_A);
            Check(Fn<TOpenSect>(admin, 10)(admin,uidPtr,IntPtr.Zero,MAPI_FORCE_ACCESS,out section), "OpenProfileSection(read)");
            uint count=0;
            Check(Fn<TGetProps>(section,5)(section,tags,0,out count,out vals),"GetProps(read)");
            if (count != 2 || vals==IntPtr.Zero) throw new InvalidOperationException("Unexpected GetProps size");
            IntPtr e=Prop(vals,count,0x3D0B), s=Prop(vals,count,0x3D0F);
            r.Entry=e!=IntPtr.Zero && Tag(e)==PR_SERVICE_ENTRY_NAME_A ? A(e) : "<other/missing>";
            bool missing=s!=IntPtr.Zero && Tag(s)==PR_SUPPORT_ERROR && Marshal.ReadInt32(s,8)==E_NOT_FOUND;
            r.Support=missing ? "MISSING" : GetMultiA(s) ?? "<other/missing>";
            bool candidate=r.Entry=="PRXServiceEntry@40" && missing;
            bool already=r.Entry=="PRXServiceEntry" && r.Support=="pstprx.dll";
            r.Status=candidate ? "CANDIDATE" : already ? "OK" : "SKIP";
            MAPIFreeBuffer(vals); vals=IntPtr.Zero;
            if (!apply || !candidate) return r;
            Release(ref section);

            // Re-open for EXCLUSIVE WRITE; failure leaves account untouched.
            Check(Fn<TOpenSect>(admin,10)(admin,uidPtr,IntPtr.Zero,MAPI_FORCE_ACCESS|MAPI_MODIFY,out section),"OpenProfileSection(write)");
            count=0;
            Check(Fn<TGetProps>(section,5)(section,tags,0,out count,out vals),"GetProps(write)");
            if (count != 2 || vals==IntPtr.Zero) throw new InvalidOperationException("Unexpected GetProps(write) size");
            e=Prop(vals,count,0x3D0B); s=Prop(vals,count,0x3D0F);
            bool same=e!=IntPtr.Zero && Tag(e)==PR_SERVICE_ENTRY_NAME_A && A(e)=="PRXServiceEntry@40"
                && s!=IntPtr.Zero && Tag(s)==PR_SUPPORT_ERROR && Marshal.ReadInt32(s,8)==E_NOT_FOUND;
            MAPIFreeBuffer(vals); vals=IntPtr.Zero;
            if (!same) throw new InvalidOperationException("Profile section has changed since scan; no modification");

            IntPtr propBuf=IntPtr.Zero, entryText=IntPtr.Zero, supportText=IntPtr.Zero, supportArray=IntPtr.Zero, problems=IntPtr.Zero;
            try {
                propBuf=Marshal.AllocHGlobal(PROP_SIZE*2);
                for(int i=0;i<PROP_SIZE*2;i++) Marshal.WriteByte(propBuf,i,0);
                entryText=Marshal.StringToHGlobalAnsi("PRXServiceEntry");
                supportText=Marshal.StringToHGlobalAnsi("pstprx.dll");
                supportArray=Marshal.AllocHGlobal(IntPtr.Size);
                Marshal.WriteIntPtr(supportArray,supportText);
                Marshal.WriteInt32(propBuf,0,unchecked((int)PR_SERVICE_ENTRY_NAME_A));
                Marshal.WriteIntPtr(propBuf,8,entryText);
                Marshal.WriteInt32(propBuf,PROP_SIZE,unchecked((int)PR_SERVICE_SUPPORT_FILES_A));
                Marshal.WriteInt32(propBuf,PROP_SIZE+8,1);  // MVszA.cValues
                Marshal.WriteIntPtr(propBuf,PROP_SIZE+16,supportArray); // MVszA.lppszA
                int hr=Fn<TSetProps>(section,8)(section,2,propBuf,out problems);
                Check(hr,"SetProps");
                if (problems!=IntPtr.Zero && Marshal.ReadInt32(problems,0)!=0)
                    throw new InvalidOperationException("SetProps reported property problems");
                Check(Fn<TSaveChanges>(section,4)(section,0),"SaveChanges");
            }
            finally {
                if(problems!=IntPtr.Zero) MAPIFreeBuffer(problems);
                if(propBuf!=IntPtr.Zero) Marshal.FreeHGlobal(propBuf);
                if(entryText!=IntPtr.Zero) Marshal.FreeHGlobal(entryText);
                if(supportText!=IntPtr.Zero) Marshal.FreeHGlobal(supportText);
                if(supportArray!=IntPtr.Zero) Marshal.FreeHGlobal(supportArray);
            }
            Release(ref section);

            // Read-back verification.
            Check(Fn<TOpenSect>(admin,10)(admin,uidPtr,IntPtr.Zero,MAPI_FORCE_ACCESS,out section),"OpenProfileSection(verify)");
            count=0;
            Check(Fn<TGetProps>(section,5)(section,tags,0,out count,out vals),"GetProps(verify)");
            e=Prop(vals,count,0x3D0B); s=Prop(vals,count,0x3D0F);
            bool verified=e!=IntPtr.Zero && Tag(e)==PR_SERVICE_ENTRY_NAME_A && A(e)=="PRXServiceEntry"
                && GetMultiA(s)=="pstprx.dll";
            if (!verified) throw new InvalidOperationException("Read-back verification failed after SaveChanges");
            r.Status="FIXED";
            r.Entry="PRXServiceEntry";
            r.Support="pstprx.dll";
            r.Detail="Saved and verified";
            return r;
        }
        catch(Exception ex) { r.Status="ERROR"; r.Detail=ex.Message; return r; }
        finally {
            if(vals!=IntPtr.Zero) MAPIFreeBuffer(vals);
            Release(ref section);
            if(tags!=IntPtr.Zero) Marshal.FreeHGlobal(tags);
            if(uidPtr!=IntPtr.Zero) Marshal.FreeHGlobal(uidPtr);
        }
    }

    private static void WalkServices(IntPtr profAdmin, string profileName, string uidFilter,
                                     bool apply, List<Result> output)
    {
        IntPtr admin=IntPtr.Zero,table=IntPtr.Zero,rows=IntPtr.Zero,cols=IntPtr.Zero;
        try {
            Check(Fn<TAdminServices>(profAdmin,11)(profAdmin,profileName,IntPtr.Zero,IntPtr.Zero,0,out admin),"AdminServices");
            Check(Fn<TGetTable>(admin,4)(admin,0,out table),"GetMsgServiceTable");
            cols=NewTags(PR_SERVICE_NAME_A,PR_DISPLAY_NAME_A,PR_SERVICE_UID);
            Check(Fn<TSetColumns>(table,7)(table,cols,0),"SetColumns(services)");
            for (;;) {
                Check(Fn<TQueryRows>(table,19)(table,32,0,out rows),"QueryRows(services)");
                if(rows==IntPtr.Zero) throw new InvalidOperationException("NULL QueryRows(services)");
                int n=Marshal.ReadInt32(rows,0);
                if(n<0 || n>32) throw new InvalidOperationException("Invalid rows count");
                for(int i=0;i<n;i++) {
                    IntPtr row=IntPtr.Add(rows,8+i*ROW_SIZE);
                    uint c=unchecked((uint)Marshal.ReadInt32(row,4));
                    IntPtr props=Marshal.ReadIntPtr(row,8);
                    IntPtr pname=Prop(props,c,0x3D09);
                    if(pname==IntPtr.Zero || Tag(pname)!=PR_SERVICE_NAME_A || A(pname)!="INTERSTOR") continue;
                    IntPtr puid=Prop(props,c,0x3D0C);
                    string uid=UidHex(puid);
                    string display="INTERSTOR";
                    IntPtr pd=Prop(props,c,0x3001);
                    if(pd!=IntPtr.Zero && Tag(pd)==PR_DISPLAY_NAME_A) display=A(pd);
                    if(uid==null) {
                        output.Add(new Result {Profile=profileName,Service=display,Status="ERROR",Detail="Invalid/missing PR_SERVICE_UID"});
                        continue;
                    }
                    if(!String.IsNullOrEmpty(uidFilter) && !String.Equals(uidFilter,uid,StringComparison.OrdinalIgnoreCase)) continue;
                    output.Add(Visit(admin,profileName,display,uid,apply));
                }
                FreeProws(rows); rows=IntPtr.Zero;
                if(n==0) break;
            }
        }
        catch(Exception ex) { output.Add(new Result {Profile=profileName,Status="ERROR",Detail=ex.Message}); }
        finally {
            if(rows!=IntPtr.Zero) FreeProws(rows);
            if(cols!=IntPtr.Zero) Marshal.FreeHGlobal(cols);
            Release(ref table);
            Release(ref admin);
        }
    }
    public static List<Result> Run(string profileFilter, string uidFilter, bool apply)
    {
        if(IntPtr.Size!=8) throw new InvalidOperationException("x64 only");
        List<Result> result=new List<Result>();
        IntPtr admin=IntPtr.Zero, table=IntPtr.Zero, rows=IntPtr.Zero, cols=IntPtr.Zero;
        Check(MAPIInitialize(IntPtr.Zero),"MAPIInitialize");
        try {
            Check(MAPIAdminProfiles(0,out admin),"MAPIAdminProfiles");
            Check(Fn<TGetTable>(admin,4)(admin,0,out table),"GetProfileTable");
            cols=NewTags(PR_DISPLAY_NAME_A);
            Check(Fn<TSetColumns>(table,7)(table,cols,0),"SetColumns(profiles)");
            for (;;) {
                Check(Fn<TQueryRows>(table,19)(table,32,0,out rows),"QueryRows(profiles)");
                if(rows==IntPtr.Zero) throw new InvalidOperationException("NULL QueryRows(profiles)");
                int n=Marshal.ReadInt32(rows,0);
                if(n<0 || n>32) throw new InvalidOperationException("Invalid profile rows count");
                for(int i=0;i<n;i++) {
                    IntPtr row=IntPtr.Add(rows,8+i*ROW_SIZE);
                    uint c=unchecked((uint)Marshal.ReadInt32(row,4));
                    IntPtr props=Marshal.ReadIntPtr(row,8);
                    IntPtr pn=Prop(props,c,0x3001);
                    if(pn==IntPtr.Zero || Tag(pn)!=PR_DISPLAY_NAME_A) continue;
                    string name=A(pn);
                    if(!String.IsNullOrEmpty(profileFilter) && name!=profileFilter) continue;
                    WalkServices(admin,name,uidFilter,apply,result);
                }
                FreeProws(rows); rows=IntPtr.Zero;
                if(n==0) break;
            }
        }
        finally {
            if(rows!=IntPtr.Zero) FreeProws(rows);
            if(cols!=IntPtr.Zero) Marshal.FreeHGlobal(cols);
            Release(ref table);
            Release(ref admin);
            MAPIUninitialize();
        }
        return result;
    }
}
'@

if (-not ('OutlookMapiInterop' -as [type])) {
    Add-Type -TypeDefinition $cs -Language CSharp -ErrorAction Stop
}

# Always scan all targets BEFORE modifying service-table contents. This also makes
# -All robust if the message service table is refreshed after SaveChanges.
$targets = @()
if ($Apply -and $All) {
    $snapshot = @([OutlookMapiInterop]::Run($Profile, $null, $false))
    if ($snapshot.Count -eq 0) {
        throw "No INTERSTOR services found in profile '$Profile'. No changes attempted."
    }
    $scanErrors = @($snapshot | Where-Object { $_.Status -eq 'ERROR' })
    if ($scanErrors.Count -gt 0) {
        $scanErrors | Select-Object Profile,Service,Status,Detail | Format-Table -AutoSize | Out-Host
        throw 'The initial scan has errors. No changes attempted.'
    }
    $targets = @($snapshot | Where-Object { $_.Status -eq 'CANDIDATE' })
    $uniqueIds = @($targets | ForEach-Object { $_.Uid } | Select-Object -Unique)
    if ($uniqueIds.Count -ne $targets.Count) {
        throw 'Duplicate INTERSTOR service UID in this profile; refusing bulk changes.'
    }
    Write-Host "Profile: $Profile" -ForegroundColor Cyan
    Write-Host "Candidate services: $($targets.Count)" -ForegroundColor Cyan
    if ($targets.Count -gt 0) {
        $targets | Select-Object Service,Uid,Entry,Support | Format-Table -AutoSize | Out-Host
    }
    else {
        Write-Host 'No candidates. Nothing to change; no backup required.' -ForegroundColor Green
        return
    }
}

if ($Apply) {
    $current = Get-Process -Id $PID
    $active = @(Get-Process -Name OUTLOOK,MFCMAPI -ErrorAction SilentlyContinue |
        Where-Object { $_.SessionId -eq $current.SessionId })
    if ($active.Count -gt 0) {
        throw 'Outlook or MFCMAPI is running in this RDS session. Close both before applying changes.'
    }

    New-Item -ItemType Directory -Path $BackupRoot -Force | Out-Null
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
    $backupPath = Join-Path $BackupRoot "OutlookProfiles-$stamp.reg"
    & reg.exe export 'HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles' $backupPath /y | Out-Null
    if ($LASTEXITCODE -ne 0 -or !(Test-Path -LiteralPath $backupPath)) {
        throw 'Registry export failed. No changes attempted.'
    }
    Write-Host "Profile registry backup: $backupPath" -ForegroundColor Yellow
}

if ($Apply -and $All) {
    $collected = New-Object System.Collections.ArrayList
    foreach ($target in $targets) {
        Write-Host "Repairing $($target.Service) [$($target.Uid)] ..."
        try {
            $resultSet = @([OutlookMapiInterop]::Run($Profile, $target.Uid, $true))
            if ($resultSet.Count -ne 1) {
                throw "Expected one MAPI service for UID $($target.Uid), got $($resultSet.Count)."
            }
            $r = $resultSet[0]
            if ($r.Uid -ne $target.Uid -or $r.Profile -ne $Profile) {
                throw "MAPI returned a different profile/service for UID $($target.Uid)."
            }
            [void]$collected.Add($r)
        }
        catch {
            [void]$collected.Add([pscustomobject]@{
                Profile=$Profile; Service=$target.Service; Uid=$target.Uid
                Status='ERROR'; Entry=$target.Entry; Support=$target.Support
                Detail=$_.Exception.Message
            })
        }
    }
    $rows = @($collected.ToArray())
}
else {
    $rows = @([OutlookMapiInterop]::Run($Profile, $Uid, [bool]$Apply))
    if ($Apply -and $rows.Count -ne 1) {
        Write-Warning "Expected one service, got $($rows.Count). Check profile and UID."
    }
}

$rows | Select-Object Profile,Service,Status,Entry,Support | Format-Table -AutoSize -Wrap
$rows | Where-Object { $_.Status -in @('CANDIDATE','FIXED','ERROR','SKIP') } |
    Select-Object Profile,Service,Uid,Status,Detail | Format-List

if ($Apply) {
    $log = Join-Path $BackupRoot ("result-" + (Get-Date -Format 'yyyyMMdd-HHmmss-fff') + '.csv')
    $rows | Export-Csv -Path $log -NoTypeInformation -Encoding UTF8
    Write-Host "Log: $log" -ForegroundColor Yellow
    if ($All) {
        $fixedCount = @($rows | Where-Object { $_.Status -eq 'FIXED' }).Count
        $skippedCount = @($rows | Where-Object { $_.Status -eq 'SKIP' -or $_.Status -eq 'OK' }).Count
        $errorCount = @($rows | Where-Object { $_.Status -eq 'ERROR' }).Count
        Write-Host "Bulk result: FIXED=$fixedCount SKIPPED=$skippedCount ERRORS=$errorCount" -ForegroundColor Cyan
    }
}
if ($rows.Count -eq 0) {
    Write-Warning 'No matching profiles/services found. Run without -Profile/-Uid for all current-user profiles.'
}
if (@($rows | Where-Object { $_.Status -eq 'ERROR' }).Count -gt 0) {
    throw 'At least one MAPI operation failed; see ERROR rows above. Previously successful changes were NOT rolled back.'
}