# Outlook-TS

Outlook-TS is a Microsoft 365 Outlook troubleshooting toolkit. The first version focuses on issues caused by moving between classic Outlook for Windows and new Outlook for Windows, then expands into Exchange Online and Microsoft Graph checks.

## What v1 Does

- Collects local Windows evidence with PowerShell:
  - New Outlook toggle registry state.
  - Classic Outlook profile registry presence.
  - Installed new Outlook AppX package.
  - Office Click-to-Run version/channel.
  - Outlook COM add-ins and risky load states.
  - Recent Outlook-related Application event log entries.
- Analyzes the collector JSON with a TypeScript CLI.
- Produces actionable findings with severity, evidence, and repair guidance.
- Includes a conservative repair script for forcing classic Outlook launch behavior.
- Collects optional cloud evidence from Exchange Online PowerShell:
  - Mailbox existence and recipient type.
  - CAS flags such as OWA, MAPI, EWS, IMAP, POP, ActiveSync.
  - Mailbox size, item counts, deleted item counts, and quota percent.
  - Archive and auto-expanding archive state.
  - Mailbox-level forwarding.
  - Optional FullAccess delegate summary.
  - Calendar folder statistics, default calendar settings, and calendar folder permissions.
  - Calendar item-count capacity risk.
- Collects optional Microsoft Graph evidence:
  - Mailbox settings availability.
  - Mail folder count, hidden folder count, and largest folders.
  - Inbox rule count and forwarding/delete/move rule summaries.
  - Calendar list, default calendar activity, sampled calendar events, recurring/cancelled event counts, and time zone mismatches.

## Why This First

Microsoft documents that when the new Outlook toggle is enabled, attempts to launch classic Outlook may redirect to new Outlook. Their documented fallback is setting `HKCU\Software\Microsoft\Office\16.0\Outlook\Preferences\UseNewOutlook` to `0`. That makes this a strong first diagnostic and repair target.

## Quick Start

For Windows machines that do not have Git or winget, use [docs/WINDOWS-SETUP.md](docs/WINDOWS-SETUP.md).

Install and build:

```bash
npm install
npm run build
```

Run the sample diagnosis:

```bash
npm run diagnose:sample
npm run diagnose:cloud-sample
```

Start the local web app:

```bash
npm run web
```

Then open:

```text
http://127.0.0.1:3000
```

On the affected Windows user profile, collect diagnostics:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\Collect-OutlookDiagnostics.ps1 -OutputPath .\outlook-diagnostics.json
```

Analyze the result:

```bash
node dist/index.js diagnose --input outlook-diagnostics.json
```

Collect Exchange Online diagnostics:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
.\scripts\Collect-ExchangeOnlineDiagnostics.ps1 -Identity user@contoso.com -Connect -IncludeMailboxPermissions
```

Collect Microsoft Graph diagnostics:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
.\scripts\Collect-GraphDiagnostics.ps1 -UserId user@contoso.com -Connect -IncludeHiddenFolders
```

For calendar sync issues, the Graph collector samples from 7 days back to 60 days forward by default:

```powershell
.\scripts\Collect-GraphDiagnostics.ps1 -UserId user@contoso.com -Connect -CalendarDaysBack 14 -CalendarDaysForward 120
```

Merge local, Exchange Online, and Graph output into one diagnostic bundle:

```powershell
.\scripts\Merge-Diagnostics.ps1 -InputPath .\outlook-diagnostics.json, .\exchange-diagnostics.json, .\graph-diagnostics.json -OutputPath .\combined.json
node .\dist\index.js diagnose --input .\combined.json --markdown .\report.md
```

You can also open the local web app and upload `combined.json`.

Preview default-calendar items older than 2 years:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
.\scripts\Remove-OldCalendarItems.ps1 -UserId user@contoso.com -Connect
```

After reviewing `old-calendar-items.csv`, delete the previewed items:

```powershell
.\scripts\Remove-OldCalendarItems.ps1 -UserId user@contoso.com -Delete
```

Create and assign an Exchange calendar retention policy for items older than 2 years:

```powershell
.\scripts\New-CalendarRetentionPolicy.ps1 `
  -Identity user@contoso.com `
  -Connect `
  -Assign `
  -StartManagedFolderAssistant
```

Force classic Outlook launch behavior for the signed-in Windows user:

```powershell
.\scripts\Repair-NewOutlookToggle.ps1 -ForceClassicLaunch
```

## Microsoft 365 Expansion Path

The local collector should stay useful without tenant admin rights. Tenant-side checks can be added in two lanes:

- Exchange Online PowerShell for mailbox, CAS, protocol, shared mailbox, policy, archive, quota, and delegation checks.
- Microsoft Graph for mailbox folder/rule/calendar/delegation observations that fit Graph permissions.

Use least-privilege permissions. For a helpdesk workflow, prefer delegated sign-in for the affected user or scoped admin consent over broad application permissions.

The initial cloud collectors use Microsoft-supported PowerShell modules instead of a custom app registration. That keeps v1 practical for techs. A dedicated app with Graph auth can come later if we want a GUI, saved tenant profiles, or packaged reports.

## Useful Official References

- [Toggle out of the new Outlook for Windows](https://support.microsoft.com/en-us/office/toggle-out-of-the-new-outlook-for-windows-ec102b39-5727-418e-ae1f-a1805434640c)
- [Run new Outlook and classic Outlook side-by-side](https://support.microsoft.com/en-us/office/run-new-outlook-and-classic-outlook-side-by-side-a624c36d-c50f-43bc-9c8b-dd17b5690ffb)
- [Classic Outlook troubleshooters](https://support.microsoft.com/en-us/windows/classic-outlook-troubleshooters-086e3d66-5404-4034-9cc5-545909dcc166)
- [Microsoft Graph permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [List mailFolders with Microsoft Graph](https://learn.microsoft.com/en-us/graph/api/user-list-mailfolders)
- [Get-EXOMailbox](https://learn.microsoft.com/en-us/powershell/module/exchange/get-exomailbox)
- [Get-EXOMailboxStatistics](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-exomailboxstatistics)
