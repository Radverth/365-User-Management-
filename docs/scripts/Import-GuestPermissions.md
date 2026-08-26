# Import-GuestPermissions.ps1

Recreates guests and their group memberships in a target tenant, from the CSV that `Export-GuestPermissions.ps1` produced.

**Changes your tenant.** Rehearse with `-WhatIf` first.

```
Guests/Import-GuestPermissions.ps1
```

| | |
|---|---|
| **Modules** | `Microsoft.Graph.Users`, `Microsoft.Graph.Groups`, `Microsoft.Graph.Identity.SignIns` |
| **Loads** | `Common/InputCsv.ps1` |
| **Graph scopes** | `User.Invite.All`, `User.ReadWrite.All`, `Group.ReadWrite.All`, `Directory.ReadWrite.All` |
| **Supports `-WhatIf`** | Yes — and still writes its log |

> **Prerequisite.** The groups must already exist in the target tenant with the **same display names**. That is what memberships are matched on; group object IDs differ between tenants and cannot be used.

---

## Parameters

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-InputPath` | string | **required** | The CSV from the export |
| `-SendInvitationMessage` | switch | off | Have Microsoft email each **newly created** guest. Off means guests are created silently |
| `-CustomInvitationMessage` | string | — | Text added to that email. No effect without `-SendInvitationMessage` |
| `-InviteRedirectUrl` | string | `https://myapplications.microsoft.com` | Where the guest lands after redeeming |
| `-LogPath` | string | `.\TenantB_GuestImport_Log.csv` | Outcome of every action |
| `-TenantId` | string | — | Tenant ID or domain to sign in against |
| `-SkipInvitations` | switch | off | Do not create missing guests; only fix up membership for guests that already exist |
| `-SkipGroupMembership` | switch | off | Only create the guests |
| `-SkipOwnership` | switch | off | Add guests as members but never as group owners |
| `-MaxGuests` | int | `0` (all) | Stop after this many guests. For a pilot run |
| `-Delimiter` | string | auto | Field separator of the input CSV |

## Examples

```powershell
# Rehearsal - writes the log, changes nothing
./Import-GuestPermissions.ps1 -InputPath ./guests.csv -WhatIf

# Pilot: five guests, silent
./Import-GuestPermissions.ps1 -InputPath ./guests.csv -MaxGuests 5

# Full run, emailing each new guest
./Import-GuestPermissions.ps1 -InputPath ./guests.csv -SendInvitationMessage `
    -CustomInvitationMessage "You are being moved to our new tenant."

# Membership only, for guests already created by another process
./Import-GuestPermissions.ps1 -InputPath ./guests.csv -SkipInvitations
```

---

## Log columns

`Timestamp`, `ExternalEmail`, `DisplayName`, `GroupName`, `Action`, `Status`, `Detail`

`Action` is `Invite`, `AddMember` or `AddOwner`. `Status`:

| Status | Meaning | Action needed |
|---|---|---|
| `Created` | Guest invited | None |
| `AlreadyExists` | Guest was already present and was reused | None |
| `Added` | Membership or ownership created | None |
| `AlreadyMember` / `AlreadyOwner` | Already correct, left alone | None |
| `GroupNotFound` | No group of that name in this tenant | **Create it and re-run** |
| `Skipped` | Not importable, ambiguous name, or a skip switch was set | Read `Detail` |
| `Failed` | Graph rejected the call | Read `Detail` |
| `WhatIf` | Rehearsal only | None |

---

## Behaviour worth knowing

**Safe to re-run — by design, not by accident.** Before writing anything it builds two indexes: every existing guest in the tenant keyed by every address they are known by, and the current members and owners of each group it touches. Anything already correct is skipped. Nothing is ever removed.

**Guests are matched on `ExternalEmail`,** never on UPN. The same person is `alice_partner.com#EXT#@tenantA.onmicrosoft.com` in one tenant and `…@tenantB.onmicrosoft.com` in the other.

**A rehearsal is a real check.** Under `-WhatIf` no object is created, but every group name is still resolved against the target tenant and the log records `GroupNotFound` for any that are missing. That is the cheapest moment to find them.

**Duplicate group names are refused, not guessed.** If two groups share a display name, memberships for it are skipped and logged.

**Dynamic groups are skipped** even when the export marked them importable, because membership there is governed by a rule.

**Owners are added as members too.** A row marked `Owner` produces both an `AddMember` and an `AddOwner` action.

---

## Common errors

**`Input CSV is missing required column(s)`** — the file did not parse into columns at all. The message names the file, the columns it *did* find, and the likely cause: an Excel workbook saved with a `.csv` extension, a file with its line breaks stripped, or simply the wrong file.

**`Could not read '<file>' as a CSV … no line breaks at all`** — the records are run together on one line and cannot be recovered. Re-export.

**`Failed` rows mentioning guest access** — adding a guest to a Microsoft 365 group requires guest access to be permitted for that group and tenant.
