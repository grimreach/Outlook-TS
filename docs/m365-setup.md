# Microsoft 365 Setup Notes

## Exchange Online PowerShell Lane

Install the Exchange Online PowerShell module on an admin workstation:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Connect-ExchangeOnline
```

Useful future collection commands:

```powershell
Get-EXOMailbox -Identity user@contoso.com -Properties DisplayName,PrimarySmtpAddress,RecipientTypeDetails,OWAMailboxPolicy,EmailAddresses
Get-EXOMailboxStatistics -Identity user@contoso.com
Get-CASMailbox -Identity user@contoso.com
```

These results can be normalized into the `exchangeOnline` section of the diagnostics JSON.

The current collector wraps these checks:

```powershell
.\scripts\Collect-ExchangeOnlineDiagnostics.ps1 -Identity user@contoso.com -Connect -IncludeMailboxPermissions
```

If already connected to Exchange Online, omit `-Connect`.

## Microsoft Graph Lane

Start with delegated permissions where possible:

- `User.Read`
- `MailboxSettings.Read`
- `Mail.ReadBasic`
- `Mail.Read`

Only request broader permissions when a rule needs them. For shared mailbox diagnostics, Graph has separate delegated shared-mail permissions such as `Mail.Read.Shared`.

Potential checks:

- Mailbox settings availability.
- Folder list including hidden folders.
- Inbox rule count and suspicious forwarding rules.
- Calendar settings and time zone.
- Shared/delegated folder visibility.

The current collector uses Microsoft Graph PowerShell:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
.\scripts\Collect-GraphDiagnostics.ps1 -UserId user@contoso.com -Connect -IncludeHiddenFolders
```

If already connected to Graph with the right scopes, omit `-Connect`.

For new Outlook calendar sync issues, collect a wider calendar window:

```powershell
.\scripts\Collect-GraphDiagnostics.ps1 -UserId user@contoso.com -Connect -IncludeHiddenFolders -CalendarDaysBack 14 -CalendarDaysForward 120
```

The Graph collector requests `Calendars.Read` along with mail and mailbox settings scopes.

## Combined Report Flow

```powershell
.\scripts\Collect-OutlookDiagnostics.ps1 -OutputPath .\outlook-diagnostics.json
.\scripts\Collect-ExchangeOnlineDiagnostics.ps1 -Identity user@contoso.com -Connect -IncludeMailboxPermissions
.\scripts\Collect-GraphDiagnostics.ps1 -UserId user@contoso.com -Connect -IncludeHiddenFolders
.\scripts\Merge-Diagnostics.ps1 -InputPath .\outlook-diagnostics.json, .\exchange-diagnostics.json, .\graph-diagnostics.json -OutputPath .\combined.json
node .\dist\index.js diagnose --input .\combined.json --markdown .\report.md
```

## App Registration Shape

For a technician-facing desktop app or CLI:

1. Register an Entra ID app as a public client.
2. Enable device code or interactive auth.
3. Use delegated permissions first.
4. Store no tokens in diagnostic bundles.
5. Redact message subjects, addresses outside the affected user, and rule target addresses unless the technician opts in.
