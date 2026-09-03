# Microsoft 365 Tenant Migration Toolkit

PowerShell scripts for migrating guest users and their group permissions from one
Microsoft 365 tenant to another, plus SharePoint site permission tooling.

Every script writes its results to CSV and the ones that change anything support
`-WhatIf`.

## Start here

**New to this, or handing it to someone who is?** Read the
user guide — a step-by-step runbook written for
administrators who know Microsoft 365 but not PowerShell. It assumes the toolkit
folder was handed to them, so it is safe to pass on as-is.

If you would rather answer questions than remember parameters:

```powershell
./Start-MigrationToolkit.ps1
```

A guided menu covering every task. It explains each step in plain language,
checks what you type, and prints the equivalent command so you can repeat a run
without it later. Anything that changes your tenant is rehearsed first: it runs
in preview, shows you the result, and only commits after you say yes.

Works on Windows, Linux and macOS. It runs in the terminal rather than in a
window, because PowerShell's windowing (WinForms and WPF) is Windows-only and
these scripts have to run on Linux too.

The rest of this README documents the underlying scripts, which you can always
run directly.

| Script | Tenant | Changes anything? |
|---|---|---|
| `Start-MigrationToolkit.ps1` | Either | Only via the task you choose |
| `Guests/Export-GuestPermissions.ps1` | A (source) | No |
| `Guests/Import-GuestPermissions.ps1` | B (target) | Yes |
| `SharePoint/Set-SiteMembersToViewers.ps1` | Either | Yes |
| `SharePoint/Get-SiteOwners.ps1` | Either | No |
| `SharePoint/Get-SiteMembers.ps1` | Either | No |
| `Exchange/Set-MailboxAddresses.ps1` | Either | Yes |
| `Reports/Get-M365UserPermissionsReport.ps1` | Either | No |
| `Setup/New-AppOnlyCertificate.ps1` | — | Local files only |

Per-script reference documentation — parameters, examples, output columns and
behaviour — is in docs/scripts/, for running the scripts
directly rather than through the menu.

`Common/` holds shared helpers the other scripts dot-source (`InputCsv.ps1` for CSV
input, `PnPConnect.ps1` for SharePoint sign-in). Keep the folder structure intact —
running a script from outside the complete folder will fail with a message telling
you so.

---

## Prerequisites

Run `./Start-MigrationToolkit.ps1` and choose **Check and install what is needed** —
it lists what is missing and offers to install it for you, into your own user
account, so no administrator rights are needed.

To do it by hand, the scripts need these specific Graph submodules rather than the
whole `Microsoft.Graph` meta-module, which is around forty modules and far slower
to install:

```powershell
Install-Module Microsoft.Graph.Users                        -Scope CurrentUser
Install-Module Microsoft.Graph.Groups                       -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.SignIns             -Scope CurrentUser  # guest invitations
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser  # guest export
Install-Module PnP.PowerShell                               -Scope CurrentUser  # SharePoint scripts
```

`Install-Module Microsoft.Graph -Scope CurrentUser` also works and covers all four
submodules, if you would rather have the lot.

All scripts sign in interactively and support MFA. The two SharePoint scripts also
accept certificate-based app-only sign-in, which is the only way to reach every site
in a tenant — see App-only access to every
site.

### Entra app registration for PnP.PowerShell

PnP.PowerShell no longer ships a shared multi-tenant app, so the SharePoint scripts
need an app registration in your own tenant. Once per tenant:

```powershell
Register-PnPEntraIDAppForInteractiveLogin `
    -ApplicationName "PnP Migration" `
    -Tenant contoso.onmicrosoft.com `
    -SharePointDelegatePermissions "AllSites.FullControl" `
    -GraphDelegatePermissions "Group.Read.All","User.Read.All"
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
`Group.Read.All` and `User.Read.All` are only there so `Get-SiteOwners.ps1` can read
Microsoft 365 group owners and resolve their names. Neither script modifies
Microsoft 365 groups, so read-only Graph scopes are enough.

Registration needs an account that can create and consent app registrations —
Application Administrator or Global Administrator.

> Use your real tenant name, e.g. `contoso.onmicrosoft.com`. Getting it wrong
> registers the app in the wrong place or fails outright. If you are unsure, the
> `SourceUserPrincipalName` column of your guest export shows it after the `#EXT#@`.

### App-only access to every site (optional)

Interactive sign-in uses **delegated** permissions: your effective access is the
app's permissions intersected with your own. Being SharePoint Administrator does
not make you a site collection administrator, so reads inside sites you do not own
are refused — that is what "Attempted to perform an unauthorized operation" means.
Granting the app more permissions changes nothing while you are signing in as
yourself.

App-only sign-in fixes that. The script authenticates as the *application*, using
application permissions, which are not limited by anyone's site membership. Register
a second app with a certificate:

```powershell
Register-PnPEntraIDApp `
    -ApplicationName "PnP Migration App-Only" `
    -Tenant contoso.onmicrosoft.com `
    -OutPath C:\Certs `
    -DeviceLogin
```

If you already have the app and only need a new certificate, skip this and see
Creating or rotating the certificate —
re-running the registration creates a second app registration rather than updating
the first.

That creates the app, generates a self-signed certificate, installs it locally and
requests `Sites.FullControl.All` (SharePoint), plus `Group.ReadWrite.All` and
`User.Read.All` (Graph). A Global Administrator must grant admin consent before it
works. Note the client ID and certificate thumbprint it prints, then:

```powershell
.\Get-SiteOwners.ps1 -AllSites `
    -TenantAdminUrl https://contoso-admin.sharepoint.com `
    -ClientId <app id> `
    -Tenant contoso.onmicrosoft.com `
    -Thumbprint <certificate thumbprint>
```

Both SharePoint scripts take `-Tenant` with either `-Thumbprint` or
`-CertificatePath` (plus `-CertificatePassword` for a .pfx). Supplying a
certificate is what switches the script to app-only — there is no separate mode
switch. Nothing prompts, so this also suits scheduled runs.

#### Creating or rotating the certificate

`Setup/New-AppOnlyCertificate.ps1` generates one on any platform:

```powershell
cd Setup
./New-AppOnlyCertificate.ps1 -CommonName "PnP Migration" -ValidYears 1 `
    -OutPath ~/certs -CertificatePassword (Read-Host -AsSecureString)
```

It writes a `.pfx` (private key, used by the scripts) and a `.cer` (public key,
uploaded to Entra), prints the thumbprint, and refuses to overwrite existing files
without `-Force`. Add `-Install` to also place it in your CurrentUser certificate
store so `-Thumbprint` works.

It uses .NET's certificate API directly rather than `New-SelfSignedCertificate`
(Windows only) or `New-PnPAzureCertificate` (which fails on some PowerShell 7.4
builds),
so it behaves the same everywhere.

Then upload the `.cer`: **Entra portal → App registrations → your app →
Certificates & secrets → Certificates → Upload certificate**.

To rotate an expiring certificate without changing the app's client ID, upload the
new `.cer` alongside the old one, switch the scripts over, then delete the old
entry.

The `.pfx` authenticates as the application. It is covered by `.gitignore`, but keep
it out of shared locations, and delete it along with the app registration once the
migration is done.

#### Windows, Linux and macOS

Both forms work on all three platforms.

`-CertificatePath` is the portable choice, and the one to use in anything shared
between machines — a `.pfx` behaves identically everywhere:

```powershell
.\Get-SiteOwners.ps1 -AllSites `
    -TenantAdminUrl https://contoso-admin.sharepoint.com `
    -ClientId <app id> `
    -Tenant contoso.onmicrosoft.com `
    -CertificatePath ./PnPMigrationAppOnly.pfx
```

Add `-CertificatePassword (Read-Host -AsSecureString)` if the .pfx is protected.

`-Thumbprint` also works everywhere, but the certificate has to be installed on
the machine. PnP resolves a thumbprint through the Windows certificate store, which
does not exist elsewhere — on Linux and macOS the scripts look it up through .NET
instead (where the CurrentUser store lives under
`~/.dotnet/corefx/cryptography/x509stores`) and hand PnP the resolved certificate.
Spaces and punctuation are stripped first, so a thumbprint copied out of the
Windows certificate dialog matches as-is.

If the thumbprint is not installed, the error says so and points at
`-CertificatePath`. To install a .pfx on Linux or macOS:

```powershell
$flags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable -bor
         [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet
$cert  = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new('./yourapp.pfx', '', $flags)
$store = [System.Security.Cryptography.X509Certificates.X509Store]::new('My','CurrentUser')
$store.Open('ReadWrite'); $store.Add($cert); $store.Dispose()
```

`Sites.FullControl.All` is full write access to every site in the tenant, held by a
certificate on disk. Treat it accordingly: it is worth deleting the app registration
once the migration is finished. If you only need the owners report and would rather
not hold that, stay interactive — group-connected sites still report their real
owners through Graph, and the rest report `TenantSiteOwner`.

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
| | | On a group-connected site this is the one that matters — see below. |
| `-ExcludeLogin` | empty | Login names or emails to leave alone. Protects service accounts. |
| `-RemoveFromMembers` | `$true` | Set `-RemoveFromMembers:$false` to copy into Visitors without removing from Members. |
| `-GuestLoginPattern` | `(#ext#\|urn:spo:guest)` | Regex identifying a guest by login name. |
| `-NoPersistedLogin` | off | Sign in afresh at every site instead of reusing the cached token. |

Users are added to Visitors **before** being removed from Members, so an
interrupted run never leaves anyone with no access. If the add to Visitors fails,
the removal is skipped and the failure is logged.

Re-running is safe: anyone already in Visitors is not re-added, and anyone already
out of Members is not touched.

#### Demoting a whole team at once

On a Microsoft 365 group-connected site, the Members SharePoint group usually holds
just two kinds of principal: the connected group's **member claim** (shown in the
admin UI as "<Site> Members") and any individuals added directly.

The claim is the one that matters. Moving that single principal to Visitors demotes
everyone in the team to read-only on the site in one step — including anyone added
to the group later — while leaving Microsoft 365 group membership alone, so Teams
chat, the group mailbox and the calendar are unaffected.

It is a group rather than a user, so it needs `-IncludeSecurityGroups`. Individuals
listed alongside it are separate: guests move by default, internal staff need
`-IncludeInternalUsers`. To clear the Members group completely:

```powershell
.\Set-SiteMembersToViewers.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Team `
    -IncludeSecurityGroups -IncludeInternalUsers -WhatIf
```

#### Microsoft 365 group-connected (Teams) sites

On a group-connected site, the people the SharePoint UI calls "members" are members
of the connected Microsoft 365 group, not of the SharePoint Members group.

**This script never modifies Microsoft 365 group membership**, because removing
someone from the group also strips their Teams chat, group mailbox and calendar
access. Group-connected sites are still processed for anyone held directly in the
SharePoint Members group, and an `Info` row in the log flags the connected group so
you can review it separately.

### Signing in once instead of once per site

Both SharePoint scripts connect to each site collection in turn. They pass
`-PersistLogin` to `Connect-PnPOnline`, which caches the delegated token under
`%LOCALAPPDATA%\.m365pnppowershell` (`$HOME/.m365pnppowershell` on Linux and macOS),
encrypted with DPAPI on Windows and the Keychain elsewhere. You sign in at the
first site and the rest reuse the token — including in a later PowerShell session,
and after a reboot.

`-PersistLogin` only exists in newer PnP.PowerShell releases. On an older module the
scripts warn once at the start and every site will prompt; `Update-Module
PnP.PowerShell` fixes that.

They also disconnect only once, at the end of the run. Calling
`Disconnect-PnPOnline` between sites drops the token context and makes the next
connection prompt again.

To force a fresh sign-in per site — connecting as different accounts, or clearing a
stale token — pass `-NoPersistedLogin`. To forget the cached token entirely:

```powershell
Disconnect-PnPOnline -ClearPersistedLogin
```

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
| `OrphanedGroup` | The site is group-connected but the group no longer exists, so it has no owners to inherit. Needs a new owner. |

Guest owners are flagged in the `IsGuest` column and counted in the summary.

Two kinds of noise are removed by default and counted in the summary: the
`SHAREPOINT\system` account, which is never a person, and the tenant-wide Global
Administrator / SharePoint Administrator role claims, which appear as site
collection administrators on most sites and are identical everywhere. Pass
`-IncludeSystemPrincipals` to keep them. Duplicate rows for the same owner on the
same site are collapsed. Each
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

**The records are run together on one line.** If the error says the file contains
no line breaks, the file has been through something that stripped its line endings
— a conversion, or copy-and-paste of the contents rather than the file itself.
Nothing can be recovered from it; re-run the export and move the file as a file.

**It is the wrong file.** The `Found:` line in the error lists the real columns.
`DisplayName, UserPrincipalName, Email, ...` means you passed the output of
`Reports/Get-M365UserPermissionsReport.ps1` rather than
`Guests/Export-GuestPermissions.ps1`.

A byte-order mark on the first column name is stripped automatically.

### "A parameter cannot be found that matches parameter name 'Interactive'"

`Register-PnPEntraIDAppForInteractiveLogin` has no `-Interactive` switch — see
Entra app registration above. Drop it.
Use `-DeviceLogin` if you need the device code flow.

### Being asked to sign in at every site

Update PnP.PowerShell — the token caching the scripts rely on needs a version that
has `-PersistLogin`:

```powershell
Update-Module PnP.PowerShell
```

The scripts warn at startup when the installed version lacks it. See
Signing in once instead of once per site.

### Owner report rows say "Attempted to perform an unauthorized operation"

Expected on any site where you are not a site collection administrator — being
SharePoint Administrator does not grant that automatically. Only the
`SiteCollectionAdmin` and `OwnersGroup` sources need site access; the other sources
still work, so those sites are not blank:

- Group-connected sites still report their real owners via
  `Microsoft365GroupOwner`, read through Graph rather than from inside the site.
  For a Teams site those *are* the owners, so nothing is actually missing.
- Other sites still report `TenantSiteOwner` from the tenant listing.

Two ways to get the SharePoint group detail as well:

- **App-only sign-in** — see
  App-only access to every site. The app
  authenticates as itself with `Sites.FullControl.All` and reaches every site.
- **Make yourself a site collection administrator** on the sites you care about
  (SharePoint admin centre, or `Set-PnPTenantSite -Owners`) and re-run.

### "The specified X509 certificate store does not exist"

An older copy of the scripts passed `-Thumbprint` straight to PnP, which resolves it
through the Windows certificate store. Current versions resolve it through .NET on
Linux and macOS instead, so update to a current copy of the toolkit. If the
certificate simply is not
installed on the machine, use `-CertificatePath` with the `.pfx` — see
Windows, Linux and macOS.

### Owner names are blank on Microsoft 365 group rows

Fixed in the script. `Get-PnPMicrosoft365GroupOwner` returns only the directory
object ID and leaves `DisplayName`, `UserPrincipalName` and `Mail` empty
— a known bug in that cmdlet — so the
script uses `Get-PnPMicrosoft365Group -IncludeOwners` instead. If a row still shows
only a GUID, its `Note` says so — grant the app `User.Read.All`.

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
