# Outlook-TS Architecture

## Goals

Outlook-TS should help a technician answer four questions quickly:

1. Is this a local client/profile/toggle problem?
2. Is this a mailbox or tenant policy problem?
3. Is this a Microsoft Graph-visible data problem?
4. What is the lowest-risk repair for this specific machine and user?

## Components

### Local Collector

`scripts/Collect-OutlookDiagnostics.ps1` runs in the affected Windows user context. It does not require tenant admin rights. It emits JSON that the TypeScript CLI can analyze.

### Repair Scripts

Repair scripts should be explicit and narrow. The first one, `scripts/Repair-NewOutlookToggle.ps1`, only sets the documented `UseNewOutlook` value to `0` for the current user unless the technician adds process-stop behavior.

### Diagnostic Engine

`src/rules.ts` contains deterministic rules. This keeps the tool explainable: every finding has evidence, a recommendation, and optionally a repair command.

### Microsoft 365 Connectors

Future collectors can add `exchangeOnline` and `graph` sections to the same JSON contract:

- Exchange Online PowerShell: mailbox existence, CAS flags, mailbox statistics, policy assignment, shared mailbox access, archive/quota.
- Microsoft Graph: mailbox settings, folder visibility, rules, calendar settings, delegated/shared mailbox access where permissions allow.

The first Microsoft 365 collectors are now PowerShell-based:

- `scripts/Collect-ExchangeOnlineDiagnostics.ps1`
- `scripts/Collect-GraphDiagnostics.ps1`
- `scripts/Merge-Diagnostics.ps1`

This lets Windows techs collect tenant-side evidence without waiting for a packaged desktop app or custom Entra ID app registration.

## Classic/New Outlook Issue Model

The first known incident pattern is:

1. User toggles from classic Outlook to new Outlook.
2. Classic Outlook launch redirects into new Outlook.
3. User toggles back and forth.
4. Local preference/profile state becomes unclear, or classic profile state is damaged.

The v1 tool checks:

- Whether new Outlook is installed.
- Whether `UseNewOutlook` is set to `1`.
- Whether classic Outlook profiles exist.
- Whether a default classic Outlook profile exists.
- Whether add-ins or recent crash events point to a separate startup problem.

## Next Rules To Add

- Detect unsupported workflows for new Outlook, especially public folder-heavy users.
- Detect mailbox CAS settings that block expected Outlook web-backed behavior.
- Compare classic Outlook profile count against default profile.
- Detect stale Autodiscover override registry keys.
- Detect add-ins that commonly break classic Outlook startup.
- Detect oversized mailbox, archive, or folder count risk from Exchange Online.
- Detect inbox rules, forwarding, or hidden folders from Graph where permissions allow.

## Current Cloud Rules

- Missing Exchange Online mailbox.
- OWA disabled.
- MAPI disabled for classic Outlook.
- EWS disabled.
- Mailbox at or above 85% and 95% quota thresholds.
- Mailbox-level forwarding.
- Many explicit FullAccess delegates.
- Graph mailbox settings unavailable.
- Many hidden folders.
- Forwarding or redirect inbox rules.
- Large inbox rule counts.
- Calendar collector permission or access errors.
- No calendars returned by Graph.
- Many calendars.
- Empty default calendar in the sampled sync window.
- Large recurring/cancelled calendar event counts.
- Calendar event time zone mismatches.
