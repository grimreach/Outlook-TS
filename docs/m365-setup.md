# Microsoft 365 Setup Notes

## Exchange Online PowerShell Lane

Install the Exchange Online PowerShell module on an admin workstation:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Connect-ExchangeOnline
```

Install the Graph submodules used by the collector:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Mail -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Calendar -Scope CurrentUser -Force
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

The Exchange Online collector also checks calendar folder statistics, default calendar folder settings, and calendar folder permissions.

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

## Calendar Capacity Cleanup

If Microsoft or diagnostics show calendar capacity pressure, first make sure archive is enabled:

```powershell
Enable-Mailbox user@contoso.com -Archive
Enable-Mailbox user@contoso.com -AutoExpandingArchive
```

Preview old default-calendar items before deleting anything:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
.\scripts\Remove-OldCalendarItems.ps1 `
  -UserId user@contoso.com `
  -Connect `
  -OlderThanYears 2 `
  -OutputPath .\old-calendar-items.json `
  -CsvPath .\old-calendar-items.csv
```

Review the CSV export. The script skips recurring items by default because deleting old occurrences from active recurring meetings can be messy. To include recurring items in the preview, add `-IncludeRecurring`.

After review, delete the previewed non-recurring items:

```powershell
.\scripts\Remove-OldCalendarItems.ps1 `
  -UserId user@contoso.com `
  -OlderThanYears 2 `
  -Delete
```

As an admin-side alternative, run a Purview targeted compliance search against the default Calendar folder. This is useful when Graph event deletion is not enough or Microsoft asks for a Purview-based purge:

```powershell
.\scripts\Invoke-PurviewCalendarPurge.ps1 `
  -Mailbox user@contoso.com `
  -Connect `
  -OlderThanYears 2
```

Review the returned search count in Purview. Then run the purge loop:

```powershell
.\scripts\Invoke-PurviewCalendarPurge.ps1 `
  -Mailbox user@contoso.com `
  -OlderThanYears 2 `
  -Purge
```

The script converts the mailbox Calendar `FolderId` into the hex `folderid:` value used by Purview KQL. It defaults to `HardDelete`; use `-PurgeType SoftDelete` if you want a softer first pass. Purview purge actions are intentionally batch-limited per mailbox, so the script refreshes the search and loops until the targeted query returns zero items.

Create a Calendar retention tag and assign a cloned policy that preserves the mailbox's existing policy tag links:

```powershell
.\scripts\New-CalendarRetentionPolicy.ps1 `
  -Identity user@contoso.com `
  -Connect `
  -OlderThanDays 730 `
  -RetentionAction DeleteAndAllowRecovery `
  -Assign `
  -StartManagedFolderAssistant
```

Use `-RetentionAction MoveToArchive` instead if the business prefers archive-first cleanup. Use `PermanentlyDelete` only after legal/compliance approval.

## App Registration Shape

For a technician-facing desktop app or CLI:

1. Register an Entra ID app as a public client.
2. Enable device code or interactive auth.
3. Use delegated permissions first.
4. Store no tokens in diagnostic bundles.
5. Redact message subjects, addresses outside the affected user, and rule target addresses unless the technician opts in.
