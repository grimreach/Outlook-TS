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

## Why This First

Microsoft documents that when the new Outlook toggle is enabled, attempts to launch classic Outlook may redirect to new Outlook. Their documented fallback is setting `HKCU\Software\Microsoft\Office\16.0\Outlook\Preferences\UseNewOutlook` to `0`. That makes this a strong first diagnostic and repair target.

## Quick Start

Install and build:

```bash
npm install
npm run build
```

Run the sample diagnosis:

```bash
npm run diagnose:sample
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

Force classic Outlook launch behavior for the signed-in Windows user:

```powershell
.\scripts\Repair-NewOutlookToggle.ps1 -ForceClassicLaunch
```

## Microsoft 365 Expansion Path

The local collector should stay useful without tenant admin rights. Tenant-side checks can be added in two lanes:

- Exchange Online PowerShell for mailbox, CAS, protocol, shared mailbox, policy, archive, quota, and delegation checks.
- Microsoft Graph for mailbox folder/rule/calendar/delegation observations that fit Graph permissions.

Use least-privilege permissions. For a helpdesk workflow, prefer delegated sign-in for the affected user or scoped admin consent over broad application permissions.

## Useful Official References

- [Toggle out of the new Outlook for Windows](https://support.microsoft.com/en-us/office/toggle-out-of-the-new-outlook-for-windows-ec102b39-5727-418e-ae1f-a1805434640c)
- [Run new Outlook and classic Outlook side-by-side](https://support.microsoft.com/en-us/office/run-new-outlook-and-classic-outlook-side-by-side-a624c36d-c50f-43bc-9c8b-dd17b5690ffb)
- [Classic Outlook troubleshooters](https://support.microsoft.com/en-us/windows/classic-outlook-troubleshooters-086e3d66-5404-4034-9cc5-545909dcc166)
- [Microsoft Graph permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [List mailFolders with Microsoft Graph](https://learn.microsoft.com/en-us/graph/api/user-list-mailfolders)
- [Get-EXOMailbox](https://learn.microsoft.com/en-us/powershell/module/exchange/get-exomailbox)
- [Get-EXOMailboxStatistics](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-exomailboxstatistics)
