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

## App Registration Shape

For a technician-facing desktop app or CLI:

1. Register an Entra ID app as a public client.
2. Enable device code or interactive auth.
3. Use delegated permissions first.
4. Store no tokens in diagnostic bundles.
5. Redact message subjects, addresses outside the affected user, and rule target addresses unless the technician opts in.
