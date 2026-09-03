# Mailbox and group addresses — commands to run

Copy-paste sheet for `Exchange/Set-MailboxAddresses.ps1`. Run every command from inside the `Exchange` folder of the toolkit, and keep the toolkit folder intact — the script loads shared code from `Common` next to it.

Your spreadsheet needs these columns, saved as CSV. Spacing and capitals do not matter.

```
User Principal Name, Hawsons Primary, PrimarySMTPAddress, Email Address
```

**Order matters.** The old tenant has to release the domain before the new tenant can verify it, and the new tenant cannot add an address on a domain it has not verified. So: run part 1 first, move the domain, then run part 2.

---

## One-time setup

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

Sign in with an account holding **Exchange Administrator** or **Recipient Administrator**. The script prompts, so nothing else is needed.

---

## Part 1 — Old tenant: move everything onto the .onmicrosoft address

```powershell
cd <toolkit folder>/Exchange
```

**1. Rehearse.** Writes the log, changes nothing. Read the log before going further.

```powershell
./Set-MailboxAddresses.ps1 -InputPath ./mailboxes.csv -Action SetPrimaryToOnMicrosoft -WhatIf
```

**2. Pilot on five.** Old address is kept as an alias, so mail still arrives.

```powershell
./Set-MailboxAddresses.ps1 -InputPath ./mailboxes.csv -Action SetPrimaryToOnMicrosoft -MaxRecipients 5
```

**3. Everyone.** Same thing, whole file. Safe to run again — anything already done is skipped.

```powershell
./Set-MailboxAddresses.ps1 -InputPath ./mailboxes.csv -Action SetPrimaryToOnMicrosoft
```

**4. Release the domain.** This removes the old address. Mail to it stops arriving from this moment, so do it when you are ready to move the domain.

```powershell
./Set-MailboxAddresses.ps1 -InputPath ./mailboxes.csv -Action SetPrimaryToOnMicrosoft -RemoveOldPrimary -WhatIf
./Set-MailboxAddresses.ps1 -InputPath ./mailboxes.csv -Action SetPrimaryToOnMicrosoft -RemoveOldPrimary
```

Then remove the domain from the old tenant in the admin centre, and add and verify it in the new one.

---

## Part 2 — New tenant: add the old address back as an alias

Nothing anyone uses today changes: the Hawsons address stays primary.

```powershell
cd <toolkit folder>/Exchange

./Set-MailboxAddresses.ps1 -InputPath ./mailboxes.csv -Action AddAliasToNewPrimary -WhatIf
./Set-MailboxAddresses.ps1 -InputPath ./mailboxes.csv -Action AddAliasToNewPrimary
```

---

## Options you may need

| Add this | When |
|---|---|
| `-Scope Groups` | Do the distribution and Microsoft 365 groups as a separate pass |
| `-Scope Mailboxes` | Do only the mailboxes |
| `-DisableEmailAddressPolicy` | The log says `PolicyManaged` |
| `-AllowAnyNewPrimary` | The new primary is deliberately not an `.onmicrosoft.com` address |
| `-LogPath ./somewhere.csv` | You want the log somewhere specific |
| `-Delimiter ";"` | The CSV was re-saved by Excel in a locale that uses semicolons |

---

## Reading the log

The log is written every time, `-WhatIf` included. Check the `Status` column.

| Status | Meaning |
|---|---|
| `PrimaryChanged` / `AliasAdded` | Done |
| `AlreadyPrimary` / `AlreadyAlias` | Already correct, nothing written |
| `WhatIf` | Rehearsal only — nothing was written |
| `DomainNotAccepted` | Add and verify the domain in this tenant first |
| `Conflict` | Somebody else already holds that address. `Detail` names them |
| `PolicyManaged` | Re-run with `-DisableEmailAddressPolicy` |
| `DirSynced` | Comes from on-premises AD — change `proxyAddresses` there |
| `Mismatch` | The recipient found has a different primary than the spreadsheet says |
| `NotFound` | Nothing there, or it is a mail user or contact, which this does not change |
| `Failed` | Exchange rejected it. `Detail` says why |

If everything comes back `DomainNotAccepted`, you are connected to the wrong tenant.
