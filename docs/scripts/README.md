# Script reference

Documentation for running each script directly, with parameters, rather than through the `Start-MigrationToolkit.ps1` menu.

The menu is a convenience. Everything it does, these scripts do — it only asks questions and passes the answers along. Use these when you want to script a run, schedule one, or pass options the menu doesn't expose.

Each has its own page in this folder, named after the script.

| Script | Purpose | Changes anything? |
|---|---|---|
| `Guests/Export-GuestPermissions.ps1` | Guests and their group memberships → CSV | No |
| `Guests/Import-GuestPermissions.ps1` | Recreate those guests in another tenant | **Yes** |
| `SharePoint/Get-SiteOwners.ps1` | SharePoint site owners → CSV | No |
| `SharePoint/Set-SiteMembersToViewers.ps1` | Demote site members to read-only | **Yes** |
| `Setup/New-AppOnlyCertificate.ps1` | Certificate for app-only sign-in | Local files only |
| `Reports/Get-M365UserPermissionsReport.ps1` | All users and their groups → CSV | No |

---

## Before you run anything

### Keep the folder together

Four of the six scripts load shared code from the `Common` folder next to them:

| Script | Loads |
|---|---|
| `Export-GuestPermissions.ps1` | `Common/InputCsv.ps1` |
| `Import-GuestPermissions.ps1` | `Common/InputCsv.ps1` |
| `Get-SiteOwners.ps1` | `Common/InputCsv.ps1`, `Common/PnPConnect.ps1` |
| `Set-SiteMembersToViewers.ps1` | `Common/InputCsv.ps1`, `Common/PnPConnect.ps1` |
| `New-AppOnlyCertificate.ps1` | nothing — fully standalone |
| `Get-M365UserPermissionsReport.ps1` | nothing — fully standalone |

They find it relative to their own location, so the layout must stay as shipped:

```
<toolkit folder>/
├── Common/
├── Guests/
├── Reports/
├── SharePoint/
└── Setup/
```

Copying a single script somewhere else will fail with a message naming the file it couldn't find. The two standalone scripts can be moved freely.

### Modules

```powershell
Install-Module Microsoft.Graph.Users                        -Scope CurrentUser
Install-Module Microsoft.Graph.Groups                       -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.SignIns             -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
Install-Module PnP.PowerShell                               -Scope CurrentUser
```

Each script declares what it needs with `#Requires`, so it refuses to start rather than failing halfway if something is absent.

### PowerShell version

PowerShell 7 or later. The scripts run on Windows, Linux and macOS.

---

## Conventions shared by all of them

**`-WhatIf` on anything that writes.** `Import-GuestPermissions.ps1` and `Set-SiteMembersToViewers.ps1` support it. Both still write their log file under `-WhatIf`, because the log is the point of a rehearsal.

**CSV in, CSV out.** Input files are read with a delimiter scan (comma, semicolon, tab, pipe), a byte-order-mark strip, and a check that rejects an Excel workbook that has been given a `.csv` name. Force a separator with `-Delimiter` if detection picks wrong.

**Output paths.** Every `-OutputPath` / `-LogPath` accepts a relative or absolute path. Parent folders are created if missing. A bare filename is written to the current directory.

**Re-running is safe.** Nothing removes a user, a member, or an owner. The write scripts read current state first and skip anything already correct.
