# Microsoft 365 Tenant Migration Toolkit

PowerShell scripts for migrating guest users and their group permissions from one
Microsoft 365 tenant to another, plus SharePoint site permission tooling.

Every script writes its results to CSV and the ones that change anything support
`-WhatIf`.

| Script | Tenant | Changes anything? |
|---|---|---|
| [`Guests/Export-GuestPermissions.ps1`](Guests/Export-GuestPermissions.ps1) | A (source) | No |
| [`Guests/Import-GuestPermissions.ps1`](Guests/Import-GuestPermissions.ps1) | B (target) | Yes |
| [`SharePoint/Set-SiteMembersToViewers.ps1`](SharePoint/Set-SiteMembersToViewers.ps1) | Either | Yes |
| [`SharePoint/Get-SiteOwners.ps1`](SharePoint/Get-SiteOwners.ps1) | Either | No |
| [`Reports/Get-M365UserPermissionsReport.ps1`](Reports/Get-M365UserPermissionsReport.ps1) | Either | No |

`Common/InputCsv.ps1` is a shared helper the other scripts dot-source. Keep the
folder structure intact — running a script from outside a full clone will fail with
a message telling you so.

---

## Prerequisites

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser      # guest scripts
Install-Module PnP.PowerShell  -Scope CurrentUser      # SharePoint scripts
```

All scripts sign in interactively and support MFA.

### Entra app registration for PnP.PowerShell

PnP.PowerShell no longer ships a shared multi-tenant app, so the SharePoint scripts
need an app registration in your own tenant. Once per tenant:

```powershell
Register-PnPEntraIDAppForInteractiveLogin `
    -ApplicationName "PnP Migration" `
    -Tenant contoso.onmicrosoft.com `
    -SharePointDelegatePermissions "AllSites.FullControl" `
    -GraphDelegatePermissions "Group.Read.All"
```

A browser opens for you to sign in and consent. Note the client ID it prints and
pass it as `-ClientId` to the SharePoint scripts.

There is **no `-Interactive` parameter** — interactive sign-in is what this cmdlet
does by default, which is what the `ForInteractiveLogin` in its name means. (The
switch exists on the older `Register-PnPAzureADApp`, which is a different cmdlet.)
If no browser is available, or MFA makes the browser flow awkward, add
`-DeviceLogin` to use the device code flow instead.

The permission parameters above are optional but worth setting. Left off, the
default grant is `AllSites.FullControl`, `Group.ReadWrite.All`,
`User.ReadWrite.All` and `TermStore.ReadWrite.All` — more than these scripts need.
`AllSites.FullControl` covers everything the SharePoint scripts do, and
`Group.Read.All` is only there so `Get-SiteOwners.ps1` can read Microsoft 365 group
owners. Neither script modifies Microsoft 365 groups, so the read-only Graph scope
is enough.

Registration needs an account that can create and consent app registrations —
Application Administrator or Global Administrator.

> Use your real tenant name, e.g. `contoso.onmicrosoft.com`. Getting it wrong
> registers the app in the wrong place or fails outright. If you are unsure, the
> `SourceUserPrincipalName` column of your guest export shows it after the `#EXT#@`.

### Graph permissions

The scripts request these scopes on connect; a Global Administrator may need to
consent the first time.

| Script | Scopes |
|---|---|
| `Export-GuestPermissions.ps1` | `User.Read.All`, `Group.Read.All`, `Directory.Read.All` |
| `Import-GuestPermissions.ps1` | `User.Invite.All`, `User.ReadWrite.All`, `Group.ReadWrite.All`, `Directory.ReadWrite.All` |

For SharePoint you need SharePoint Administrator, or site collection administrator
rights on each site you touch.

---

## Guest migration

Groups must already exist in tenant B **with the same display names** — that is what
the import matches on. Object IDs differ between tenants and cannot be used.

### 1. Export from tenant A

```powershell
cd Guests
.\Export-GuestPermissions.ps1
```

Produces `TenantA_GuestPermissions.csv` (one row per guest per group) and
`TenantA_GuestPermissions_Unsupported.csv` (memberships Graph cannot recreate).

Useful switches:

| Switch | Effect |
|---|---|
| `-IncludeDisabledGuests` | Keep guests blocked from sign-in. Excluded by default. |
| `-IncludePendingAcceptance` | Keep guests who never redeemed their invitation. Excluded by default. |
| `-ExcludeGroupOwnership` | Skip the owned-groups lookup. Faster on large tenants. |

Only **direct** memberships are exported. Nested memberships are inherited
automatically in tenant B once the direct membership exists — exporting them would
create incorrect direct memberships.

### 2. Rehearse against tenant B

```powershell
.\Import-GuestPermissions.ps1 -InputPath .\TenantA_GuestPermissions.csv -WhatIf
```

Nothing is written to the tenant, but the log CSV is still produced. Every group
name is resolved against tenant B, so this is how you find out which groups are
missing or ambiguously named **before** you commit. Check the log for
`GroupNotFound` rows and fix those first.

### 3. Import

```powershell
# Pilot: five guests, no email sent
.\Import-GuestPermissions.ps1 -InputPath .\TenantA_GuestPermissions.csv -MaxGuests 5

# Full run, silent
.\Import-GuestPermissions.ps1 -InputPath .\TenantA_GuestPermissions.csv

# Full run, emailing each new guest
.\Import-GuestPermissions.ps1 -InputPath .\TenantA_GuestPermissions.csv -SendInvitationMessage
```

#### The invitation email

Guests are created **silently by default**. They appear in the directory and can be
added to groups immediately, but receive nothing.

Add `-SendInvitationMessage` to have Microsoft email each newly invited guest, and
`-CustomInvitationMessage "..."` to add your own text to that email. Only guests
created on that run are emailed — re-running never re-mails anyone.

#### Other switches

| Switch | Effect |
|---|---|
| `-MaxGuests <n>` | Stop after n guests. For pilot runs. |
| `-SkipInvitations` | Only fix up group membership for guests that already exist. |
| `-SkipGroupMembership` | Only create the guests. |
| `-SkipOwnership` | Add guests as members but never as group owners. |
| `-InviteRedirectUrl` | Where the guest lands after redeeming. Defaults to the My Apps portal. |
| `-Delimiter` | Field separator of the input CSV. Detected automatically when omitted. |

### Re-running is safe

The import is additive and never removes anything. On a second run:

- A guest whose external address already exists is reused, not re-invited.
- Existing group members and owners are detected first and left alone.
- Nothing is removed from any group, ever.

Verified against a mocked tenant: a first run performed 2 invitations, 2 member
adds and 1 owner add; an immediate second run performed **zero** write operations
and logged every row as `AlreadyExists` / `AlreadyMember` / `AlreadyOwner`.

Guests are matched across tenants by their **external (home tenant) email address**,
taken from `mail`, then `otherMails`, then decoded from the UPN. A guest's UPN is
tenant-specific — `alice@partner.com` is
`alice_partner.com#EXT#@tenanta.onmicrosoft.com` in A and
`alice_partner.com#EXT#@tenantb.onmicrosoft.com` in B — so it cannot be used for
matching.

### What the import cannot do

These are flagged `Importable = False` by the export and written to the unsupported
CSV for manual handling:

| Group type | Why | What to do |
|---|---|---|
| Distribution lists | Graph cannot add members to them | `Add-DistributionGroupMember` in Exchange Online PowerShell |
| Mail-enabled security groups | Same | Same |
| Dynamic groups | Membership comes from a rule, not member objects | Recreate the rule in tenant B |

The import also logs, rather than guesses, when a group name does not exist in
tenant B or when two groups share a display name.

---

## SharePoint

### Demote site members to viewers

Moves people out of a site's **Members** group (Edit) into its **Visitors** group
(Read).

```powershell
cd SharePoint

# Rehearse first
.\Set-SiteMembersToViewers.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Project -ClientId $id -WhatIf

# Several sites from a CSV with a SiteUrl column
.\Set-SiteMembersToViewers.ps1 -SitesCsvPath .\sites.csv -ClientId $id
```

**Guests only by default.** Internal staff are left alone unless you ask for them:

| Parameter | Default | Effect |
|---|---|---|
| `-IncludeInternalUsers` | off | Also demote internal (non-guest) users. |
| `-IncludeSecurityGroups` | off | Also demote security groups sitting in the Members group. Reported otherwise. |
| `-ExcludeLogin` | empty | Login names or emails to leave alone. Protects service accounts. |
| `-RemoveFromMembers` | `$true` | Set `-RemoveFromMembers:$false` to copy into Visitors without removing from Members. |
| `-GuestLoginPattern` | `(#ext#\|urn:spo:guest)` | Regex identifying a guest by login name. |

Users are added to Visitors **before** being removed from Members, so an
interrupted run never leaves anyone with no access. If the add to Visitors fails,
the removal is skipped and the failure is logged.

Re-running is safe: anyone already in Visitors is not re-added, and anyone already
out of Members is not touched.

#### Microsoft 365 group-connected (Teams) sites

On a group-connected site, the people the SharePoint UI calls "members" are members
of the connected Microsoft 365 group, not of the SharePoint Members group.

**This script never modifies Microsoft 365 group membership**, because removing
someone from the group also strips their Teams chat, group mailbox and calendar
access. Group-connected sites are still processed for anyone held directly in the
SharePoint Members group, and an `Info` row in the log flags the connected group so
you can review it separately.

### Report site owners

Read-only.

```powershell
# Every site in the tenant - the usual starting point
.\Get-SiteOwners.ps1 -AllSites -TenantAdminUrl https://contoso-admin.sharepoint.com -ClientId $id

# Specific sites
.\Get-SiteOwners.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Project -ClientId $id

# From a CSV with a SiteUrl column
.\Get-SiteOwners.ps1 -SitesCsvPath .\sites.csv -ClientId $id
```

Your tenant admin URL is your SharePoint host with `-admin` inserted:
`https://contoso.sharepoint.com` becomes `https://contoso-admin.sharepoint.com`.

Running the script with no arguments prints these three forms rather than doing
anything.

"Owner" means three different things in SharePoint, and all three are collected —
each owner is one row tagged with its `OwnerSource`:

| Source | What it is |
|---|---|
| `SiteCollectionAdmin` | Full control of the site collection. The real administrators, and the ones most often missed. |
| `OwnersGroup` | Members of the site's associated Owners SharePoint group. |
| `Microsoft365GroupOwner` | Owners of the connected Microsoft 365 group. These are site owners even when the Owners group looks empty. |
| `TenantSiteOwner` | The site collection's primary owner from the tenant listing. `-AllSites` only. |

Guest owners are flagged in the `IsGuest` column and counted in the summary. Each
source can be turned off with `-IncludeSiteCollectionAdmins:$false`,
`-IncludeOwnersGroup:$false` or `-IncludeMicrosoft365GroupOwners:$false`.

With `-AllSites`, OneDrive personal sites are excluded unless you add
`-IncludeOneDrive`.

`-AllSites` opens each site individually to read its groups, so allow time on a
large tenant. Two things follow from that:

- A fourth source, `TenantSiteOwner`, appears — the primary owner recorded against
  the site collection, taken from the tenant listing rather than from inside the
  site.
- Being SharePoint Administrator does **not** make you a site collection
  administrator everywhere. Sites that refuse to open still get their
  `TenantSiteOwner` row, alongside an `Error` row saying group-level owners are
  missing for that site. Nothing is silently dropped — filter the CSV on
  `OwnerSource = Error` to see which sites need you to grant yourself access and
  re-run.

---

## Troubleshooting

### "Input CSV is missing required column(s)" — and every column is listed

The columns are almost certainly present. The file did not parse into columns at
all, so all of them look missing. The error names the file and the columns it
actually found, which tells you which case you are in:

**The file is an Excel workbook wearing a `.csv` name.** The most common cause.
Opening a CSV in Excel and using *File > Save As* writes a real `.xlsx` while
leaving the extension alone. The scripts detect this and say so explicitly. Fix it
with *File > Save As > CSV UTF-8 (Comma delimited) (\*.csv)*, or re-run the export
and feed the file straight in without opening it in Excel.

**The separator is not a comma.** A file saved in a locale that uses semicolons
parses as a single column. The scripts try `,` `;` tab and `|` automatically and
report which one they used, so this now loads on its own. Force one with
`-Delimiter ";"` if detection picks wrong.

**It is the wrong file.** The `Found:` line in the error lists the real columns.
`DisplayName, UserPrincipalName, Email, ...` means you passed the output of
`Reports/Get-M365UserPermissionsReport.ps1` rather than
`Guests/Export-GuestPermissions.ps1`.

A byte-order mark on the first column name is stripped automatically.

### "A parameter cannot be found that matches parameter name 'Interactive'"

`Register-PnPEntraIDAppForInteractiveLogin` has no `-Interactive` switch — see
[Entra app registration](#entra-app-registration-for-pnppowershell) above. Drop it.
Use `-DeviceLogin` if you need the device code flow.

### The export returned fewer guests than expected

Disabled guests and guests who never redeemed their original invitation are both
excluded by default. Add `-IncludeDisabledGuests` and `-IncludePendingAcceptance`
to keep them.

---

## Suggested migration order

1. `Get-SiteOwners.ps1` against tenant A — record who owns what before anything moves.
2. `Export-GuestPermissions.ps1` against tenant A.
3. Confirm the groups exist in tenant B with matching display names.
4. `Import-GuestPermissions.ps1 -WhatIf` against tenant B; resolve every `GroupNotFound`.
5. `Import-GuestPermissions.ps1 -MaxGuests 5` — pilot, and confirm the results by hand.
6. `Import-GuestPermissions.ps1` — full run, adding `-SendInvitationMessage` when you are ready for guests to be told.
7. Work through the unsupported CSV in Exchange Online PowerShell.
8. `Set-SiteMembersToViewers.ps1 -WhatIf`, then for real.
