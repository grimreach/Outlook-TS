# Graph App Setup

Outlook-TS can run read-only Microsoft Graph diagnostics from the local web UI. This avoids installing Graph PowerShell modules on every technician workstation.

## Entra App Registration

Create an app registration in Microsoft Entra admin center:

1. Go to App registrations.
2. Create a new registration named `Outlook-TS Local Diagnostics`.
3. Supported account type: single tenant.
4. Redirect URI platform: Single-page application.
5. Redirect URI:

```text
http://127.0.0.1:3000/
```

If technicians use another local port, add that redirect URI too.

## API Permissions

Add delegated Microsoft Graph permissions:

```text
User.Read
MailboxSettings.Read
Mail.ReadBasic
Mail.Read
Calendars.Read
```

Grant admin consent so technicians are not prompted for each permission.

## Run Locally

Start Outlook-TS:

```powershell
cd C:\Users\Remote\Outlook-TS
npm.cmd run web
```

Open:

```text
http://127.0.0.1:3000/
```

In Graph Diagnostics:

1. Enter the tenant ID or tenant domain.
2. Enter the application client ID.
3. Enter the target user.
4. Select Sign In.
5. Select Collect Graph.

The tool writes a Graph diagnostics JSON bundle into the report input and analyzes it with the same rules as uploaded diagnostics.

## Current Scope

The browser Graph lane is read-only. It collects:

- Mailbox settings.
- Mail folder counts.
- Inbox rule summaries.
- Calendar list.
- Calendar sync-window activity.
- Calendar time zone mismatch signals.

Bulk delete and retention actions remain in the Purview and Exchange Online PowerShell lane.
