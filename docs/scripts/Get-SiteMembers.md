# Get-SiteMembers.ps1

Reports who has member-level access to SharePoint sites, to CSV.

**Read-only.** Nothing is modified.

The companion to `Get-SiteOwners.ps1`: that one answers *who runs this site*, this one answers *who can get into it*.

---

## Where access comes from

A report that only reads the Members group misses most of it. All of these are collected, each row tagged with its source:

| `MemberSource` | What it is |
|---|---|
| `MembersGroup` | The site's associated Members group |
| `VisitorsGroup` | The associated Visitors group — read-only access, but access |
| `SharePointGroup` | Any other SharePoint group on the site. Real sites accumulate these and they are easily forgotten |
| `DirectPermission` | People and security groups given permission on the site itself rather than through a group |
| `Microsoft365GroupMember` | Members of the connected Microsoft 365 group, for group-connected (Teams) sites. These people have access even when the Members group looks empty |
| `OwnersGroup` | The associated Owners group. Only with `-IncludeOwnersGroup` |
| `OrphanedGroup` | The site is connected to a Microsoft 365 group that no longer exists |
| `Error` | Something could not be read. The `Note` says what and why |
| `None` | Nothing was found for this site |

Each person produces one row per route, so somebody in the Members group who also holds a direct permission appears twice. That is the point: it is what you have to unpick to remove their access.

Owners are left to `Get-SiteOwners.ps1`, which also covers site collection administrators — this script does not look at those.

---

## Parameters

| Parameter | Type | Default | Meaning |
|---|---|---|---|
| `-SiteUrl` | string[] | — | One or more site collection URLs |
| `-SitesCsvPath` | string | — | CSV with a `SiteUrl` column, instead of `-SiteUrl` |
| `-AllSites` | switch | off | Every site in the tenant. Needs `-TenantAdminUrl` |
| `-TenantAdminUrl` | string | — | e.g. `https://contoso-admin.sharepoint.com`. Required with `-AllSites` |
| `-IncludeOneDrive` | switch | off | With `-AllSites`, also include personal OneDrive sites |
| `-ClientId` | string | — | Client ID of the Entra app registration PnP signs in with |
| `-IncludeMembersGroup` | bool | `$true` | The associated Members group |
| `-IncludeVisitorsGroup` | bool | `$true` | The associated Visitors group |
| `-IncludeOtherGroups` | bool | `$true` | Every other SharePoint group on the site |
| `-IncludeDirectPermissions` | bool | `$true` | Principals given permission on the site directly |
| `-IncludeMicrosoft365GroupMembers` | bool | `$true` | Members of the connected Microsoft 365 group |
| `-IncludeOwnersGroup` | switch | off | Also include the Owners group, for one file covering everybody |
| `-GuestsOnly` | switch | off | Report only guests |
| `-GuestLoginPattern` | string | `(#ext#\|urn:spo:guest)` | Regex identifying a guest login |
| `-IncludeSystemPrincipals` | switch | off | Keep `SHAREPOINT\system`, the Everyone claims and the tenant admin role claims |
| `-OutputPath` | string | `.\SharePoint_SiteMembers.csv` | Destination CSV |
| `-Delimiter` | string | auto | Field separator of `-SitesCsvPath` |
| `-Tenant`, `-Thumbprint`, `-CertificatePath`, `-CertificatePassword` | | — | App-only sign-in. Same as `Get-SiteOwners.ps1` |
| `-NoPersistedLogin` | switch | off | Sign in afresh for every site instead of reusing a cached token |

## Examples

```powershell
# One site
./Get-SiteMembers.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Project -ClientId $id

# Every site in the tenant
./Get-SiteMembers.ps1 -AllSites -TenantAdminUrl https://contoso-admin.sharepoint.com -ClientId $id

# Every guest with access to any site, and how they got it
./Get-SiteMembers.ps1 -AllSites -TenantAdminUrl https://contoso-admin.sharepoint.com -ClientId $id -GuestsOnly

# From a site list
./Get-SiteMembers.ps1 -SitesCsvPath ./sites.csv -ClientId $id -OutputPath ./members.csv
```

---

## Output columns

| Column | Meaning |
|---|---|
| `SiteUrl` | The site |
| `SiteTitle` | Its title |
| `Template` | Its template |
| `IsGroupConnected` | Whether it is connected to a Microsoft 365 group |
| `MemberSource` | Which route the access came by — see the table above |
| `GroupName` | The group the person is in, where the route is a group |
| `MemberName` | Display name |
| `MemberLogin` | Login/claim, which is what identifies them uniquely |
| `MemberEmail` | Email address, where SharePoint has one |
| `PrincipalType` | `User`, `SecurityGroup` or `SharePointGroup` |
| `Roles` | The permission levels behind this row |
| `IsGuest` | Whether the login matches `-GuestLoginPattern` |
| `Note` | Anything worth saying about the row |

---

## Behaviour worth knowing

**The `Roles` column is read from the site, not assumed.** One pass over the site's role assignments does two jobs: it produces the `DirectPermission` rows, and it records what each SharePoint group is actually allowed to do. So a custom group that has been given Full Control is reported as Full Control rather than as ordinary membership — which is the finding you were looking for.

**Limited Access is not reported.** SharePoint grants it automatically so somebody can reach a single item in a library. Reporting it as site access is misleading, so entries whose only permission is `Limited Access` or `Web-Only Limited Access` are dropped.

**System and tenant-wide claims are filtered by default,** because they appear on nearly every site and bury the real names: `SHAREPOINT\system`, `Everyone`, `Everyone except external users`, and the Global/SharePoint Administrator role claims. The number removed is reported at the end.

Run once with `-IncludeSystemPrincipals` all the same. A site carrying `Everyone except external users` is open far wider than its group membership suggests — that is a finding, not noise.

**`-GuestsOnly` is the version people actually read.** Before a tenant migration the question is which external people can reach which sites, and the full report is mostly internal noise. Rows that carry no principal — `Error`, `None`, `OrphanedGroup` — are never filtered away, so a site that could not be read still says so.

**A site that refuses to open produces an `Error` row, not silence.** Being SharePoint Administrator does not make you a site collection administrator everywhere. Under interactive sign-in expect some of these on a tenant-wide run; app-only with `Sites.FullControl.All` reaches everything.

**Microsoft 365 group members sometimes come back as bare object IDs.** The PnP cmdlet does not always populate display names. Rather than emit a blank line, the row falls back to the directory object ID and the `Note` says to grant the app `User.Read.All` to resolve names.

**Sign-in happens once, not per site.** The connection is reused across sites, so a tenant-wide run does not prompt repeatedly.

---

## Requirements

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

An app registration for PnP to sign in with — `Register the app in Entra ID` in the toolkit menu creates one. For tenant-wide coverage, app-only sign-in with a certificate and `Sites.FullControl.All`.

Keep the toolkit folder intact: this script loads `Common/InputCsv.ps1` and `Common/PnPConnect.ps1` from beside it.
