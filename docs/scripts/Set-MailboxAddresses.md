# Set-MailboxAddresses.ps1

Rewrites mailbox email addresses from a spreadsheet, for the two moves a cross-tenant migration needs: releasing a vanity address in the old tenant, and adding it as an alias in the new one.

**Changes your tenant.** Supports `-WhatIf`.

Two jobs, one spreadsheet, chosen with `-Action`. They run against different tenants and are normally done days or weeks apart.

| `-Action` | Tenant | What it does |
|---|---|---|
| `SetPrimaryToOnMicrosoft` | old | Finds each mailbox by its `PrimarySMTPAddress` and makes the `.onmicrosoft` address from the *Email Address* column the primary instead |
| `AddAliasToNewPrimary` | new | Finds each mailbox by the *Hawsons Primary* address and adds the `PrimarySMTPAddress` to it as an alias, leaving the primary alone |

> This one is not on the `Start-MigrationToolkit.ps1` menu — run it directly.

---

## The spreadsheet

Save it as CSV. Four columns are understood, and **spacing and capitals do not matter** — `PrimarySMTPAddress`, `Primary SMTP Address` and `primary smtp address` are all the same column. Extra columns are ignored.

| Column | Also accepted as | Needed by |
|---|---|---|
| User Principal Name | `UserPrincipalName`, `User Principle Name`, `UPN`, `Identity` | Optional in both. Used to find the mailbox if the address lookup fails, and shown in the log |
| Hawsons Primary | `HawsonsPrimary`, `New Primary`, `Target Primary` | `AddAliasToNewPrimary` |
| PrimarySMTPAddress | `Primary SMTP Address`, `PrimarySMTP`, `Primary` | both |
| Email Address | `EmailAddress`, `OnMicrosoftAddress`, `Email` | `SetPrimaryToOnMicrosoft` |

Only the columns the chosen action actually reads are required. Ask for `AddAliasToNewPrimary` without a *Hawsons Primary* column and the script names the missing column, lists the headers it did find, and stops before connecting to anything.

---

## Parameters

| Parameter | Type | Default | Meaning |
|---|---|---|---|
| `-InputPath` | string | **required** | The spreadsheet, as CSV |
| `-Action` | string | **required** | `SetPrimaryToOnMicrosoft` or `AddAliasToNewPrimary` |
| `-RemoveOldPrimary` | switch | off | `SetPrimaryToOnMicrosoft` only. Remove the old primary instead of keeping it as an alias |
| `-AllowAnyNewPrimary` | switch | off | Accept a new primary that is not an `.onmicrosoft.com` address |
| `-DisableEmailAddressPolicy` | switch | off | Turn off the email address policy on mailboxes that have one, so their addresses can be set |
| `-LogPath` | string | `.\Mailbox_<action>_Log.csv` | Outcome of every row |
| `-MaxMailboxes` | int | `0` (all) | Stop after this many rows. For a pilot run |
| `-AdminUpn` | string | — | Sign in as this account. Prompts if omitted |
| `-Organization` | string | — | Tenant domain, for app-only sign-in |
| `-AppId` | string | — | Application (client) ID, for app-only sign-in |
| `-CertificateThumbprint` | string | — | Certificate thumbprint, for app-only sign-in |
| `-Delimiter` | string | auto | Field separator of the input CSV |

The three app-only parameters go together — pass one and the script asks for all three. Omit all three to sign in interactively.

## Examples

```powershell
# Rehearsal against the old tenant. Writes the log, changes nothing
./Set-MailboxAddresses.ps1 -InputPath ./mailboxes.csv -Action SetPrimaryToOnMicrosoft -WhatIf

# Pilot: five mailboxes, old address kept as an alias on each
./Set-MailboxAddresses.ps1 -InputPath ./mailboxes.csv -Action SetPrimaryToOnMicrosoft -MaxMailboxes 5

# The real release: .onmicrosoft becomes primary and the old address is removed,
# so the domain can be verified in the new tenant
./Set-MailboxAddresses.ps1 -InputPath ./mailboxes.csv -Action SetPrimaryToOnMicrosoft -RemoveOldPrimary

# Against the new tenant: add each old address as an alias. No primary changes
./Set-MailboxAddresses.ps1 -InputPath ./mailboxes.csv -Action AddAliasToNewPrimary
```

---

## Output columns

| Column | Meaning |
|---|---|
| `Timestamp` | When the row was processed |
| `Action` | Which of the two jobs was running |
| `UserPrincipalName` | From the spreadsheet, for matching rows back |
| `Mailbox` | The mailbox that was found, by its primary address at the time |
| `Address` | The address this row was about |
| `Status` | See below |
| `PrimaryBefore` | The mailbox's primary address before the change |
| `PrimaryAfter` | And after it |
| `Detail` | Plain-language reason, and what to do about it |

### Status values

| Status | Meaning |
|---|---|
| `PrimaryChanged` | Done. The `.onmicrosoft` address is now the primary |
| `AliasAdded` | Done. The old address is on the mailbox as an alias and the primary is unchanged |
| `AlreadyPrimary` | Nothing to do — the address is already the primary |
| `AlreadyAlias` | Nothing to do — the alias is already on the mailbox |
| `Mismatch` | A mailbox was found, but its primary is not what the spreadsheet says. Left alone |
| `NotFound` | No mailbox for that address. If the address belongs to something that is not a mailbox, the `Detail` says what it is |
| `Conflict` | The alias is already in use by somebody else. `Detail` names them |
| `DomainNotAccepted` | The address is on a domain this tenant has not accepted. Add and verify the domain first |
| `PolicyManaged` | An email address policy manages the mailbox. Re-run with `-DisableEmailAddressPolicy` |
| `DirSynced` | The mailbox comes from on-premises Active Directory. Change its `proxyAddresses` there |
| `Refused` | The new primary is not an `.onmicrosoft.com` address. Re-run with `-AllowAnyNewPrimary` if that is deliberate |
| `Skipped` | The row was unusable — a blank or malformed address. `Detail` says which |
| `WhatIf` | Rehearsal only. Nothing was written |
| `Failed` | Exchange rejected the change, or the result was not what was asked for. `Detail` carries the message |

---

## Behaviour worth knowing

**Nothing but SMTP addresses is touched.** `SIP:`, `X500:` and `EUM:` entries are carried across exactly as they are. Those are what keep free/busy lookups and replies to old messages working, and losing them is a support call per user.

**The old primary is kept by default.** `SetPrimaryToOnMicrosoft` demotes the old address to an alias, so mail sent to it still arrives. That is deliberately not enough to release the domain: a domain cannot be verified in the new tenant while addresses in the old one still use it. `-RemoveOldPrimary` is the step that actually frees it, and it is the one thing here that loses something — mail to the old address stops being delivered the moment it goes.

Running with `-RemoveOldPrimary` after a run without it does the right thing: the primary is already correct, so only the removal is applied.

**The primary is set by rewriting the whole address list,** not by adding to it. Only a full array can say "this one is primary and the previous one is not" in a single write. The list is rebuilt with the new primary as `SMTP:` and everything else as `smtp:`, duplicates removed case-insensitively.

**`AddAliasToNewPrimary` adds with a lowercase `smtp:` prefix,** which is what makes an address an alias. An uppercase prefix would silently take over as the primary, which is the opposite of what this action is for. After each write the mailbox is read back and the primary is checked; if it moved, the row is logged `Failed` rather than reported as a success.

**Safe to re-run.** A mailbox already in the wanted state is reported and skipped without a write. Both actions were verified to perform zero writes on a second run.

**Checks happen before writes, not instead of error handling.** Accepted domains are read once at the start; directory sync, address policies and address conflicts are checked per mailbox. Each produces a named status and a sentence saying what to do, because the equivalent Exchange error says none of that.

**A mailbox is found by address first, then by user principal name.** `Get-Mailbox` matches on any of a mailbox's addresses, so the lookup still works after a run has changed the primary.

---

## Common errors

**`Input CSV is missing the column(s) -Action ... needs`** — the header wording is not one the script recognises. It lists the headers it found; rename the one you meant in the spreadsheet and save as CSV again. Spacing and capitals are already ignored.

**`The e-mail address policy is enabled on this mailbox`** — logged as `PolicyManaged` rather than thrown. Re-run with `-DisableEmailAddressPolicy`.

**`The operation couldn't be performed because the object is being synchronized from your on-premises organization`** — logged as `DirSynced`. The addresses live in Active Directory; change `proxyAddresses` there and let sync carry it up.

**Everything comes back `DomainNotAccepted`** — you are connected to the wrong tenant. The summary line printed at sign-in says how many accepted domains were found; check that against the tenant you meant.

**Everything comes back `Mismatch`** — the spreadsheet's `PrimarySMTPAddress` column no longer reflects the tenant, usually because a run already changed them. The `PrimaryBefore` column in the log shows what each mailbox actually has.

---

## Requirements

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

The signed-in account needs a role that can change mailbox addresses — **Exchange Administrator** or **Recipient Administrator**. Global Administrator also works but is more than the job needs.
