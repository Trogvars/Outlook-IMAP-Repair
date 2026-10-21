# Reports of Outlook error 0x80040605 when opening IMAP/OST settings

[English README](README.md) · [Русский README](README.ru.md)

This page collects **independent public reports** resembling the symptom fixed by [Outlook-IMAP-Repair](https://github.com/Trogvars/Outlook-IMAP-Repair): classic Outlook shows **"An unknown error occurred, error code: 0x80040605"** when opening **Data Files → Settings** for an IMAP OST file.

> These reports demonstrate the **same symptom**. They do **not** prove that the users' MAPI properties match the configuration we observed. The fix must be preceded by a read-only scan. This is a community troubleshooting project, not an official Microsoft patch.

## Closest matching discussions

| Year | Source | Why it is relevant | Reply limitations |
| --- | --- | --- | --- |
| 2023 | [Microsoft Q&A — Can't compact Outlook 2021 OST file](https://learn.microsoft.com/en-us/answers/questions/4619759/cant-compact-outlook-2021-ost-file) | Outlook 2021 + IMAP + OST + error under Data Files → Settings; user did not want to recreate four accounts. | Migrated **locked** thread. |
| 2020 | [Microsoft Q&A — błąd 0x80040605 outlook 365](https://learn.microsoft.com/pl-pl/answers/questions/4554000/b-d-0x80040605-outlook-365) | OST settings/compacting failed after Windows and Office updates, per the author. | Migrated **locked** thread. |
| 2020 | [Ten Forums — Outlook 2019 backup/restoration](https://www.tenforums.com/microsoft-office-365/150682-help-me-understand-hot-mess-outlook-2019-backup-restoration.html) | Outlook 2019, IMAP-only mailboxes, multiple OST files failing on opening Settings. | Current reply permissions **not verified**. |
| 2018 | [Microsoft Q&A — Outlook 2016 (Portuguese)](https://learn.microsoft.com/pt-br/answers/questions/4509733/erro-desconhecido-0x80040605-outlook-2016) | Same OST settings error; author said switching Office from 32-bit to 64-bit helped. | **Different workaround**; cannot assume our MAPI change is required. |
| 2015 | [Microsoft Q&A — foutcode 0x80040605 (Dutch)](https://learn.microsoft.com/nl-nl/answers/questions/4507914/foutcode-0x80040605) | Error opening Outlook 2016 Data Files settings for an IMAP mailbox. | Historical migrated thread. |

## Related technical articles

- [TheWindowsClub — Fix 0x80040605 Unknown error in Outlook](https://www.thewindowsclub.com/fix-0x80040605-unknown-error-in-outlook) lists ScanPST, profile recreation, registry export, and Office repair, but not the two MAPI property changes.
- [Microsoft MFCMAPI](https://github.com/microsoft/mfcmapi) can be used to inspect the failing **INTERSTOR** message service.

## Observed repair (Outlook LTSC 2021 x64 + IMAP)

The tool checks for **both** conditions in a pre-existing **INTERSTOR** service:

1. **PR_SERVICE_ENTRY_NAME_A** — tag 0x3D0B001E, type PT_STRING8 — has value **PRXServiceEntry@40**.
2. **PR_SERVICE_SUPPORT_FILES_A** — tag 0x3D0F101E, type PT_MV_STRING8 — is **missing**; MFCMAPI displays tag 0x3D0F000A, PT_ERROR, HRESULT 0x8004010F.

If both match, it sets:

- **PR_SERVICE_ENTRY_NAME_A** = **PRXServiceEntry**
- **PR_SERVICE_SUPPORT_FILES_A** = ANSI multi-value string array with a single element **pstprx.dll**

A single-mailbox repair was verified with Extended MAPI and read-back verification, without recreating the Outlook profile or OST file. This is a **limited observed fix**, not a general cure for every occurrence of 0x80040605.

Keep Outlook closed, inspect a read-only scan, back up the profile, and validate one mailbox before using bulk mode.

## How to try it

Download the scripts from the [main repository](https://github.com/Trogvars/Outlook-IMAP-Repair).

~~~bat
REM Read-only scan from CMD:
Run-OutlookIMAP.cmd

REM After confirming candidates and backing up, repair one selected profile:
Run-OutlookIMAP.cmd -Apply -Profile "Outlook" -All
~~~

Run in the affected Windows user's session. The launcher selects Windows PowerShell 5.1 x64. This tool is independent and is **not supported or endorsed by Microsoft**.

## Sharing with others

Describe this as an **independently observed workaround**, link to the repository, and ask readers to confirm the specific MAPI properties in a read-only scan before modifying anything. Do not post real mailbox addresses, UIDs, or Outlook profile exports. Some Microsoft Q&A threads were locked after migration and do not accept new replies.
