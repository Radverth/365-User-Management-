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

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-OutputPath` | string | `.\M365_User_Permissions_Report.csv` | Where to write the report |
| `-TenantId` | string | — | Tenant ID or domain to sign in against. Useful when your account exists in more than one tenant |

## Examples

```powershell
# Current folder, default name
./Get-M365UserPermissionsReport.ps1

# Named output, specific tenant
./Get-M365UserPermissionsReport.ps1 -OutputPath C:\Audit\permissions.csv `
    -TenantId contoso.onmicrosoft.com
```

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
