# Set-SiteMembersToViewers.ps1

Moves principals out of a site's **Members** group and into its **Visitors** group, demoting them from Edit to Read.

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

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-SiteUrl` | string[] | **required**¹ | One or more site collection URLs |
| `-SitesCsvPath` | string | **required**¹ | CSV with a `SiteUrl` column |
| `-IncludeInternalUsers` | switch | off | Also demote internal staff. **Off means guests only** |
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

¹ `-SiteUrl` and `-SitesCsvPath` are alternatives; supply exactly one.

## Examples

```powershell
# Rehearse - guests only, one site
./Set-SiteMembersToViewers.ps1 -ClientId <app id> `
    -SiteUrl https://contoso.sharepoint.com/sites/Project -WhatIf

# Everyone, including the connected Team, protecting one account
./Set-SiteMembersToViewers.ps1 -ClientId <app id> `
    -SiteUrl https://contoso.sharepoint.com/sites/Project `
    -IncludeInternalUsers -IncludeSecurityGroups `
    -ExcludeLogin svc-admin@contoso.com

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

`Timestamp`, `SiteUrl`, `PrincipalName`, `LoginName`, `Email`, `PrincipalType`, `IsGuest`, `Status`, `Detail`

| Status | Meaning |
|---|---|
| `AddedToVisitors` | Granted Read |
| `RemovedFromMembers` | Edit removed. The demotion is complete |
| `AlreadyVisitor` | Was already in Visitors, not re-added |
| `KeptInMembers` | `-RemoveFromMembers:$false` was set |
| `Excluded` | Matched `-ExcludeLogin` |
| `Skipped` | Internal user, or a group principal, and the matching switch was not set |
| `Info` | Group-connected site notice |
| `SiteSkipped` / `SiteFailed` | The whole site could not be processed |
| `Failed` | The move failed — read `Detail` |
| `WhatIf` | Rehearsal only |

---

## Behaviour worth knowing

**Order is deliberate.** Each principal is added to Visitors *before* being removed from Members. If the add fails, the removal is skipped and logged — nobody is left with no access to the site.

**Guests only, by default.** Internal staff are untouched unless `-IncludeInternalUsers` is set. A guest is identified by login name matching `-GuestLoginPattern`, which covers both B2B guests (`#ext#`) and SharePoint-only external users (`urn:spo:guest`).

**On a Teams site, `-IncludeSecurityGroups` is the one that matters.** The Members group of a group-connected site holds the connected group's *member claim*, shown in the admin centre as **"&lt;Site&gt; Members"**. Moving that single principal to Visitors makes the whole team read-only on the site at once, including people who join the group later.

Microsoft 365 group membership itself is never modified — deliberately. Removing someone there would also strip their Teams chat, group mailbox and calendar. The script flags group-connected sites in the log with an `Info` row so you can review them.

Guest-versus-internal only applies to actual users. A group principal is neither, so `-IncludeSecurityGroups` alone is enough to move it — it does not also need `-IncludeInternalUsers`.

**Safe to re-run.** Anyone already in Visitors is not re-added; anyone already out of Members is untouched.

**Sites with no Visitors group are skipped, not half-done.** Removing people from Members with nowhere to put them would silently revoke access, so the whole site is skipped and logged.

**One sign-in for the whole run.** The connection is reused across sites and disconnected once at the end.
