# Set-SiteMembersToViewers.ps1

Demotes people from Edit to Read on a SharePoint site, in both places access comes from: membership of the site's **Members** group, and any permission granted to them **directly** on the site. With `-IncludeOtherGroups`, every other SharePoint group on the site as well.

Both matter — moving someone out of Members changes nothing if they also hold Edit directly.

**Changes your tenant.** Rehearse with `-WhatIf` first.

```
SharePoint/Set-SiteMembersToViewers.ps1
```

| | |
|---|---|
| **Modules** | `PnP.PowerShell` |
| **Loads** | `Common/InputCsv.ps1`, `Common/PnPConnect.ps1` |
| **Needs** | An Entra app registration; site collection administrator rights, or app-only sign-in |
| **Supports `-WhatIf`** | Yes — and still writes its log |

---

## Parameters

### Choosing sites — three mutually exclusive forms

| Parameter set | Parameters |
|---|---|
| One or more URLs | `-SiteUrl <string[]>` |
| From a spreadsheet | `-SitesCsvPath <file>` — needs a `SiteUrl` column |
| Whole tenant | `-AllSites -TenantAdminUrl <url>` (both required together) |

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-SiteUrl` | string[] | — | One or more site collection URLs |
| `-SitesCsvPath` | string | — | CSV with a `SiteUrl` column |
| `-AllSites` | switch | — | Every site in the tenant |
| `-TenantAdminUrl` | string | — | Required with `-AllSites`, e.g. `https://contoso-admin.sharepoint.com` |
| `-IncludeOneDrive` | switch | off | With `-AllSites`, also include personal OneDrive sites |
| `-Scope` | `Guests` / `Staff` / `Both` | `Guests` | Who to demote. Applies to group membership and direct permissions alike |
| `-IncludeOtherGroups` | switch | off | Also empty every other SharePoint group on the site into Visitors — see below |
| `-SkipDirectPermissions` | switch | off | Only fix group membership, leaving direct site permissions untouched |
| `-IncludeInternalUsers` | switch | off | Superseded by `-Scope Both` and equivalent to it. Still accepted so existing commands work |
| `-IncludeSecurityGroups` | switch | off | Also demote group principals inside Members — see below |
| `-GuestLoginPattern` | string | `(#ext#\|urn:spo:guest)` | Regex identifying a guest by login name |
| `-ExcludeLogin` | string[] | empty | Login names or emails to leave alone entirely |
| `-RemoveFromMembers` | bool | `$true` | `:$false` copies into Visitors without removing from Members |
| `-ClientId` | string | — | Entra app client ID |
| `-Tenant` | string | — | Tenant name. Required for app-only sign-in |
| `-Thumbprint` | string | — | Certificate thumbprint. Supplying it switches to app-only |
| `-CertificatePath` | string | — | `.pfx` instead of a thumbprint. Also switches to app-only |
| `-CertificatePassword` | SecureString | — | Password for the `.pfx` |
| `-NoPersistedLogin` | switch | off | Sign in afresh at each site |
| `-OutputPath` | string | `.\SharePoint_MembersToViewers_Log.csv` | Outcome of every principal considered |
| `-Delimiter` | string | auto | Field separator of `-SitesCsvPath` |

## Examples

```powershell
# Rehearse - guests only, one site
./Set-SiteMembersToViewers.ps1 -ClientId <app id> `
    -SiteUrl https://contoso.sharepoint.com/sites/Project -WhatIf

# Staff only, leaving guests as they are
./Set-SiteMembersToViewers.ps1 -ClientId <app id> `
    -SiteUrl https://contoso.sharepoint.com/sites/Project -Scope Staff

# Everyone, including the connected Team, protecting one account
./Set-SiteMembersToViewers.ps1 -ClientId <app id> `
    -SiteUrl https://contoso.sharepoint.com/sites/Project `
    -Scope Both -IncludeSecurityGroups `
    -ExcludeLogin svc-admin@contoso.com

# Group membership only, leaving direct grants alone
./Set-SiteMembersToViewers.ps1 -ClientId <app id> `
    -SiteUrl https://contoso.sharepoint.com/sites/Project -Scope Both -SkipDirectPermissions

# Every site in the tenant - rehearse this one first
./Set-SiteMembersToViewers.ps1 -AllSites `
    -TenantAdminUrl https://contoso-admin.sharepoint.com `
    -ClientId <app id> -WhatIf

# Many sites from a spreadsheet, app-only
./Set-SiteMembersToViewers.ps1 -SitesCsvPath ./sites.csv `
    -ClientId <app id> -Tenant contoso.onmicrosoft.com `
    -CertificatePath ./PnPMigration.pfx `
    -CertificatePassword (Read-Host -AsSecureString)

# Staged rollout: grant Read without removing Edit yet
./Set-SiteMembersToViewers.ps1 -ClientId <app id> `
    -SiteUrl https://contoso.sharepoint.com/sites/Project -RemoveFromMembers:$false
```

---

## Log columns

`Timestamp`, `SiteUrl`, `PrincipalName`, `LoginName`, `Email`, `PrincipalType`, `IsGuest`, `SourceGroup`, `Status`, `Detail`

`SourceGroup` names the group the person was taken out of. With `-IncludeOtherGroups` that is the column to read.

| Status | Meaning |
|---|---|
| `AddedToVisitors` | Granted Read |
| `RemovedFromMembers` | Edit removed. The demotion is complete |
| `AlreadyVisitor` | Was already in Visitors, not re-added |
| `DirectPermissionReduced` | A permission held directly on the site was reduced to Read. `Detail` names what was removed |
| `AlreadyReadOnly` | Held a direct permission that was already read-only |
| `KeptInMembers` | `-RemoveFromMembers:$false` was set |
| `Excluded` | Matched `-ExcludeLogin` |
| `Skipped` | Internal user, or a group principal, and the matching switch was not set |
| `Info` | Group-connected site notice |
| `GroupSkipped` | A group was empty, so there was nothing to move |
| `SiteSkipped` / `SiteFailed` | The whole site could not be processed |
| `Failed` | The move failed — read `Detail` |
| `WhatIf` | Rehearsal only |

---

## Behaviour worth knowing

**Order is deliberate.** Each principal is added to Visitors *before* being removed from their group. If the add fails, the removal is skipped and logged — nobody is left with no access to the site.

**`-IncludeOtherGroups` catches the second members group.** Sites that have been through a few hands often carry more than one member-type group — a plain **Members** beside the site's own **&lt;Site&gt; Members** — usually left behind by whoever set the site up. Demoting only the associated group leaves that other one exactly as it was, and the people in it carry on editing.

With the switch, every SharePoint group on the site is emptied into Visitors, not just the associated one. Three rules keep it safe:

- **The Owners group is never touched**, whatever it has been renamed to. It is found by association, not by name.
- **The Visitors group is never a source** — it is where everyone is going.
- **An empty group is logged `GroupSkipped`** rather than passed over silently.

It is off by default because a custom group can carry Full Control, and emptying one of those is a bigger change than demoting the Members group. Rehearse with `-WhatIf` and read the `SourceGroup` column before committing.

Everything else applies unchanged: `-Scope` still decides who is in scope, `-ExcludeLogin` still protects individuals, and people are still added to Visitors before being removed.

**Guests only, by default.** `-Scope Staff` does the opposite, leaving guests alone; `-Scope Both` covers everyone. A guest is identified by login name matching `-GuestLoginPattern`, which covers both B2B guests (`#ext#`) and SharePoint-only external users (`urn:spo:guest`). The same test decides scope for group membership and for direct permissions, so the two cannot disagree.

**Direct permissions.** Role assignments on the site itself are read, and anything above read-only is removed and replaced with Read. Three things are never touched:

- **`Limited Access`** — system-managed, and what lets someone reach a parent so an item-level grant works. Removing it breaks that access.
- **The site's own SharePoint groups** — those are the membership pass's business.
- **`SHAREPOINT\system`** — never a person, and removing its access breaks the site.

Permission levels already meaning read-only (`Read`, `View Only`, `Restricted View`, `Restricted Read`) are left as they are and logged as `AlreadyReadOnly`. Custom levels are treated as edit-capable, so a custom read-only level would be replaced with `Read` — check the rehearsal if you use one.

**On a Teams site, `-IncludeSecurityGroups` is the one that matters.** The Members group of a group-connected site holds the connected group's *member claim*, shown in the admin centre as **"&lt;Site&gt; Members"**. Moving that single principal to Visitors makes the whole team read-only on the site at once, including people who join the group later.

Microsoft 365 group membership itself is never modified — deliberately. Removing someone there would also strip their Teams chat, group mailbox and calendar. The script flags group-connected sites in the log with an `Info` row so you can review them.

Guest-versus-internal only applies to actual users. A group principal is neither, so `-IncludeSecurityGroups` alone is enough to move it — it does not also need `-IncludeInternalUsers`.

**Safe to re-run.** Anyone already in Visitors is not re-added; anyone already out of Members is untouched.

**Sites with no Visitors group are skipped, not half-done.** Removing people from Members with nowhere to put them would silently revoke access, so the whole site is skipped and logged.

**`-AllSites` is tenant-wide and destructive.** It enumerates every site from the
tenant admin site and demotes members across all of them. Personal OneDrive sites
are excluded unless `-IncludeOneDrive` is given. The script prints the site count
and a warning before starting. Rehearse with `-WhatIf` and read the log first —
this is the one option where a mistake is expensive to undo.

Interactive sign-in only reaches sites you administer, so a tenant-wide run will
report `SiteFailed` on the rest. App-only sign-in with `Sites.FullControl.All`
reaches all of them.

**One sign-in for the whole run.** The connection is reused across sites and disconnected once at the end.
