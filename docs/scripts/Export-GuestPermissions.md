# Export-GuestPermissions.ps1

Reads every guest user in a tenant, with the groups they belong to and own, and writes them to CSV.

**Changes nothing.** Read-only against the directory.

```
Guests/Export-GuestPermissions.ps1
```

| | |
|---|---|
| **Modules** | `Microsoft.Graph.Users`, `Microsoft.Graph.Groups`, `Microsoft.Graph.Identity.DirectoryManagement` |
| **Loads** | `Common/InputCsv.ps1` |
| **Graph scopes** | `User.Read.All`, `Group.Read.All`, `Directory.Read.All` |
| **Supports `-WhatIf`** | No — it writes nothing to the tenant |

---

## Parameters

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-OutputPath` | string | `.\TenantA_GuestPermissions.csv` | The report. This is what `Import-GuestPermissions.ps1` reads. |
| `-UnsupportedOutputPath` | string | `.\TenantA_GuestPermissions_Unsupported.csv` | Only the memberships Graph cannot recreate. Written only if there are any. |
| `-TenantId` | string | — | Tenant ID or domain to sign in against. Useful when your account exists in more than one tenant. |
| `-IncludeDisabledGuests` | switch | off | Include guests blocked from signing in. |
| `-IncludePendingAcceptance` | switch | off | Include guests who never redeemed their original invitation. |
| `-ExcludeGroupOwnership` | switch | off | Skip the owned-groups lookup. One fewer Graph call per guest — noticeably faster on a large tenant, at the cost of not recording who owns what. |

## Examples

```powershell
# Straightforward run
./Export-GuestPermissions.ps1

# Everything, including disabled and never-redeemed guests
./Export-GuestPermissions.ps1 -IncludeDisabledGuests -IncludePendingAcceptance

# Named output, specific tenant
./Export-GuestPermissions.ps1 -TenantId contoso.onmicrosoft.com `
    -OutputPath C:\Migration\guests.csv
```

---

## Output columns

| Column | Meaning |
|---|---|
| `GuestDisplayName` | Display name in the source tenant |
| `ExternalEmail` | **The identity key.** The guest's home-tenant address — the only value that survives a move between tenants |
| `SourceUserPrincipalName` | UPN in the source tenant, e.g. `alice_partner.com#EXT#@contoso.onmicrosoft.com` |
| `SourceObjectId` | Object ID in the source tenant. Recorded for audit; not used by the import |
| `ExternalUserState` | `Accepted` or `PendingAcceptance` |
| `AccountEnabled` | Whether sign-in is permitted |
| `CompanyName`, `JobTitle`, `Department` | Directory attributes, where set |
| `GroupDisplayName` | **The match key for the import.** `None` when the guest belongs to no groups |
| `GroupMailNickname` | Mail nickname, where the group has one |
| `SourceGroupId` | Group's object ID in the source tenant. Audit only — IDs differ between tenants |
| `GroupType` | `Microsoft 365 Group`, `Security group`, `Distribution list`, `Mail-enabled security group` |
| `MembershipType` | `Member`, `Owner` (both), `OwnerOnly`, or `None` |
| `Importable` | `True` if the import can recreate it |
| `SkipReason` | Why not, when `Importable` is `False` |

A guest with no groups still gets one row, so the import knows to invite them.

---

## Behaviour worth knowing

**Direct memberships only.** Nested memberships are excluded deliberately. They are re-inherited in the target tenant once the direct membership exists, so exporting them would create incorrect direct memberships.

**`ExternalEmail` is resolved in three steps** — `mail`, then `otherMails`, then decoded from the UPN by taking everything before `#EXT#` and converting the *last* underscore back to `@`. A guest with none of these is skipped and named in a warning.

**What `Importable = False` means.** These are exported and listed separately, but need handling by hand:

| Type | Why | Do this instead |
|---|---|---|
| Distribution list | Graph cannot add members | `Add-DistributionGroupMember` (Exchange Online) |
| Mail-enabled security group | Same | Same |
| Dynamic group | Membership comes from a rule | Recreate the rule in the target tenant |

**Group details are cached.** A group shared by two hundred guests is read once, not two hundred times.

**The output is verified.** After writing, the file is read back and checked for the expected row count and columns. If it cannot be parsed you get a warning naming the byte count and line-break count, and the summary marks the report as failed verification. A report that cannot be read is worse than no report.

---

## Common errors

**`No guest users found in this tenant.`** — Exactly that. The script exits cleanly.

**`No guests remain after filtering.`** — All your guests are disabled or never redeemed. Add `-IncludeDisabledGuests` and/or `-IncludePendingAcceptance`.

**Warnings about guests with no resolvable external email.** — They cannot be re-invited elsewhere, so they are omitted from the report. The warning names them.
