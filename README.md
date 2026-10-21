# Outlook IMAP Repair

**Repair existing classic Outlook IMAP profiles affected by error 0x80040605 — without recreating accounts or OST files.**

**Languages:** English | [Русский](README.ru.md)

> **Status:** pilot administrator utility. Targeted repair has been verified on Outlook LTSC 2021 x64. Bulk and RDS automation require testing before production rollout. The scripts are now included in this repository; test their behavior in your own environment before broad rollout.

## Symptoms

On some existing **classic Microsoft Outlook** profiles using IMAP, opening **File → Account Settings → Account Settings → Data Files → Settings** fails with **0x80040605**. The same error can occur in [MFCMAPI](https://github.com/microsoft/mfcmapi) when configuring the associated message service.

Mail synchronization may otherwise work. SCANPST, Office repair, and rebuilding an OST file may not fix the settings dialog. A newly created Outlook profile can work correctly, but rebuilding many IMAP accounts on an RDS server is disruptive.

The observed issue is associated with properties of certain existing **INTERSTOR** MAPI services. The exact Office update responsible has **not** been determined.

## Verified configuration change

Both of these changes were made together in successful repairs:

| MAPI property | Affected profile | Working profile |
| --- | --- | --- |
| **PR_SERVICE_ENTRY_NAME_A** — 0x3D0B001E, PT_STRING8 | PRXServiceEntry@40 | PRXServiceEntry |
| **PR_SERVICE_SUPPORT_FILES_A** — 0x3D0F101E, PT_MV_STRING8 | Missing | ANSI string array with one entry: pstprx.dll |

MFCMAPI may display the missing support-files property as:

~~~text
Property: PR_SERVICE_SUPPORT_FILES
Tag:      0x3D0F000A
Type:     PT_ERROR
Error:    0x8004010F (MAPI_E_NOT_FOUND)
~~~

This error is a **read result indicating that the property is absent** — not a real stored PT_ERROR property that should be renamed. The correct property must be **created with type PT_MV_STRING8**.

Since both changes were applied together, it has **not** been established whether either single change would be sufficient in isolation.

## Requirements

- Windows, classic **Outlook LTSC 2021 x64**, and IMAP accounts (tested combination).
- **Windows PowerShell 5.1 x64** (Desktop edition), not PowerShell 7 or 32-bit PowerShell.
- Execute as the **affected signed-in Windows user**, never as SYSTEM or another administrator.
- Close Outlook and MFCMAPI before applying changes.
- No Visual Studio installation or separate PowerShell modules required on the server. The tool relies on PowerShell Add-Type and Extended MAPI from the Outlook installation.

Other Office versions, Exchange/POP3 profiles, and other architectures are **not validated**.

Check your PowerShell process:

~~~powershell
[PSCustomObject]@{
    Version = $PSVersionTable.PSVersion.ToString()
    Edition = $PSVersionTable.PSEdition
    Is64Bit = [Environment]::Is64BitProcess
    Path    = (Get-Process -Id $PID).Path
}
~~~

The 64-bit executable is normally:

~~~text
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
~~~

When calling from a **32-bit process**, use Windows' **Sysnative** alias to access the 64-bit executable.

## Universal startup from CMD, PowerShell 5/7, x86 or x64

Use **[Run-OutlookIMAP.cmd](Run-OutlookIMAP.cmd)** as the preferred entry point. It resolves its own folder, selects 64-bit Windows PowerShell 5.1 using **Sysnative** when called from a 32-bit process (otherwise **System32**), and forwards arguments and exit codes to the main script.

~~~bat
REM Read-only diagnosis
Run-OutlookIMAP.cmd

REM Read-only diagnosis of a specific profile
Run-OutlookIMAP.cmd -Profile "Outlook"

REM Repair all matching services in one profile (close Outlook first)
Run-OutlookIMAP.cmd -Apply -Profile "Outlook" -All
~~~

From **Windows PowerShell 5.1**, **PowerShell 7** or a PowerShell x86 session:

~~~powershell
& ".\Run-OutlookIMAP.cmd" -Profile "Outlook"
& ".\Run-OutlookIMAP.cmd" -Apply -Profile "Outlook" -All
~~~

Alternatively, call the companion **[Start-OutlookIMAP.ps1](Start-OutlookIMAP.ps1)** from any Windows PowerShell / pwsh process; it performs the same x64 bootstrap:

~~~powershell
.\Start-OutlookIMAP.ps1 -Profile "Outlook"
.\Start-OutlookIMAP.ps1 -Apply -Profile "Outlook" -All
~~~

**Do not call the core Repair-OutlookIMAP-v2.ps1 directly from PowerShell 7 or x86 PowerShell.** Its architecture guard will reject those hosts. The launchers only guarantee the **correct process architecture and PowerShell edition**; they cannot override Group Policy execution restrictions, missing Outlook/MAPI installation, unavailable user profiles, or running Outlook sessions.

For per-user RDS logon repair, use [Run-RDS-OutlookRepair.cmd](Run-RDS-OutlookRepair.cmd) instead. It starts the RDS wrapper and applies changes to the signed-in user's **default Outlook profile** only.

---

## Usage

The core script is [Repair-OutlookIMAP-v2.ps1](Repair-OutlookIMAP-v2.ps1). Run these commands from the directory containing the downloaded script.

### Read-only scan

~~~powershell
.\Repair-OutlookIMAP-v2.ps1
.\Repair-OutlookIMAP-v2.ps1 -Profile "Outlook"
~~~

Possible statuses:

| Status | Meaning |
| --- | --- |
| **OK** | Both repaired values are present |
| **CANDIDATE** | Exact known-bad pair detected |
| **SKIP** | Other configuration; no changes |
| **ERROR** | Inspection or MAPI operation failed |
| **FIXED** | Change saved and properties read back successfully |

### Repair one service

Use the actual **32-character hexadecimal UID** from scan output:

~~~powershell
.\Repair-OutlookIMAP-v2.ps1 -Apply -Profile "Outlook" -Uid "0123456789ABCDEF0123456789ABCDEF"
~~~

### Repair all candidates in one profile

~~~powershell
.\Repair-OutlookIMAP-v2.ps1 -Apply -Profile "Outlook" -All
~~~

The tool only writes when **both** the old entry name and missing support-files property match. It does not rewrite unrelated services.

### RDS / terminal server

The optional [Invoke-RDS-OutlookRepair.ps1](Invoke-RDS-OutlookRepair.ps1) wrapper runs the core tool separately for selected profiles of the **current user**:

~~~powershell
# Scan the current user's default Outlook profile
.\Invoke-RDS-OutlookRepair.ps1

# Repair the default profile
.\Invoke-RDS-OutlookRepair.ps1 -Apply

# Repair all Outlook profiles of the current user (including test copies)
.\Invoke-RDS-OutlookRepair.ps1 -Apply -AllProfiles
~~~

For server-wide deployment, run in each user's context at logon (for example, with a **user** GPO logon script). A single elevated/SYSTEM task cannot correctly substitute for every user's MAPI session. If Outlook or MFCMAPI is running, the launcher defers the repair in that session. Multiple simultaneous sessions for one account need additional operational checks.

The included [Run-RDS-OutlookRepair.cmd](Run-RDS-OutlookRepair.cmd) selects **x64** Windows PowerShell via **Sysnative** when needed and runs **-Apply** for the current user's **default profile only**. Use **-AllProfiles** with the PowerShell launcher only when you intentionally want to process every profile, including test copies.

## Backups and safety

Before writing, the core script exports the current user's Outlook profile registry data:

~~~text
HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles
~~~

Backups and repair logs are saved by default to:

~~~text
%LOCALAPPDATA%\OutlookImapRepair\
~~~

- Start with **read-only scan**, then **one mailbox**, then **one profile**.
- Confirm Outlook can open OST settings, compact the file, and synchronize IMAP normally after repair.
- Bulk repair is **not transactional**: if one service fails, successfully modified services are **not automatically rolled back**.
- A .reg export is **not** a complete mailbox backup or guaranteed one-click rollback. Keep separate backups as required.
- Profile backups may contain sensitive account settings. Restrict access and **redact logs** before sharing them.
- Run the scripts only if you trust the source and have reviewed the code; do not modify active profiles.

## Implementation

The script uses **Extended MAPI** to enumerate the signed-in user's Outlook profiles and **INTERSTOR** services. It checks both original properties, writes the corrected typed values via **SetProps / SaveChanges**, and reads the properties back. It does not perform blind replacement inside binary registry values.

This project is independent and **not affiliated with Microsoft**.

## References

- [MFCMAPI](https://github.com/microsoft/mfcmapi)
- [Microsoft Learn — MAPI reference](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/mapi-reference)

## Contributing / license

Reproducible test results and contributions are welcome. Do not include real email addresses, usernames, service UIDs, or exported Outlook profiles in public reports.

A project license has not yet been chosen; add a LICENSE file before presenting the code as permissively licensed open-source software.
