# Get-M365UserPermissionsReport.ps1

Lists every user in a tenant with the groups they belong to.

**Changes nothing.** Read-only.

```
Reports/Get-M365UserPermissionsReport.ps1
```

| | |
|---|---|
| **Modules** | `Microsoft.Graph.Users`, `Microsoft.Graph.Groups` |
| **Loads** | Nothing — fully standalone, so it can be copied out of the folder |
| **Graph scopes** | `User.Read.All`, `Group.Read.All`, `Directory.Read.All` |
| **Supports `-WhatIf`** | No — it writes nothing to the tenant |

---

## Parameters

**None.** This script takes no arguments. The output path is set at the top of the file:

```powershell
$ExportPath = "M365_User_Permissions_Report.csv"
```

Edit that line to change where the report goes. It is written to the current directory by default.

```powershell
./Get-M365UserPermissionsReport.ps1
```

> This is the oldest script here and predates the others' parameter conventions. It is kept because it works and covers a different question — *all* users, not just guests. If you want it parameterised like the rest, that is a small change worth asking for.

---

## Output columns

| Column | Meaning |
|---|---|
| `DisplayName` | User's display name |
| `UserPrincipalName` | UPN |
| `Email` | Mail attribute, where set |
| `UserType` | `Member` or `Guest` |
| `AccountStatus` | `Enabled` or `Disabled` |
| `GroupName` | Group display name, or `None` |
| `GroupId` | Group object ID |
| `GroupType` | `Microsoft 365 Group / Team`, `Security Group`, `Distribution Group`, or `Unknown Group Type` |

One row per user per group. A user in no groups still gets a row, with `None` in the group columns.

---

## How it differs from Export-GuestPermissions

| | This script | `Export-GuestPermissions.ps1` |
|---|---|---|
| **Covers** | All users | Guests only |
| **Memberships** | Transitive — includes nested | Direct only |
| **Ownership** | Not recorded | Recorded |
| **Purpose** | An audit snapshot | Input to a migration |

The transitive difference matters. This report shows *effective* access, including membership inherited through nested groups — right for an audit, wrong for replaying into another tenant, which is why the migration export deliberately records direct memberships only.

Use this one for a permissions review. Use the export when you intend to import.

---

## Behaviour worth knowing

**Group details are cached.** Each group is read from Graph once, however many users belong to it.

**A user whose memberships cannot be read is reported and skipped,** not fatal — the run continues.

**Groups that cannot be read** still produce a row, falling back to the display name from the membership record, or the group ID, with `Unknown Group Type`.
