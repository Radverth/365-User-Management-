# Tenant Migration Runbook

**Moving guest users and their group permissions between Microsoft 365 tenants, and resetting SharePoint site access.** Written for administrators who know Microsoft 365 but not PowerShell.

| | |
|---|---|
| **You need** | Global Administrator in both tenants |
| **Setup time** | About 30 minutes, once |
| **Typing required** | Answering menu questions |
| **Reversible?** | Nothing is deleted, ever |

**Contents**

1. How this works
2. Part 1 — One-time setup
3. Part 2 — Moving the guests
4. Part 3 — SharePoint access
5. Part 4 — Reading the reports
6. Part 5 — When something goes wrong
7. Part 6 — Afterwards

---

## How this works

You never write PowerShell. One script asks questions and runs the right thing based on your answers:

```powershell
./Start-MigrationToolkit.ps1
```

Two rules hold throughout, and they're worth trusting.

**Nothing is ever removed.** The guest scripts only add — people, memberships, ownerships. If something already exists, it's left alone. Running any task twice changes nothing the second time.

**Every change is rehearsed.** Tasks that touch your tenant run in preview first, write a log of exactly what they *would* do, and wait for you to say yes.

Throughout this guide, tasks are marked either **[Read-only]** or **[Changes your tenant]**.

---

## Part 1 — One-time setup

*Do this once per tenant.*

### Step 1 — Install PowerShell 7

Not "Windows PowerShell" — that's the older blue one. You want **PowerShell 7**, a separate free install from Microsoft that runs on Windows, macOS and Linux.

**Check it worked.** Open a terminal, type `pwsh`, then `$PSVersionTable.PSVersion`. You want 7 or higher.

### Step 2 — Put the toolkit somewhere

You'll be given a folder containing the scripts. Put it wherever you like — your home folder is fine — and keep it intact.

**Keep the folder structure.** The scripts share code in a `Common` sub-folder. Moving individual scripts out on their own will break them, and they'll tell you so if you do.

Open a terminal in that folder before running anything:

```powershell
cd /path/to/the/toolkit/folder
```

### Step 3 — Install the modules

Start the menu and choose **Check and install what is needed** — the first option, under *First-time setup*. It lists what's missing and offers to install it for your user account only — no admin rights needed.

```powershell
./Start-MigrationToolkit.ps1
```

**What you should see.** Five modules listed, each showing either a version number or `MISSING`. Say yes to the install and it fetches them from the PowerShell Gallery. Takes a few minutes.

### Step 4 — Register an Entra app

The SharePoint tooling needs an app registration in your tenant — Microsoft removed the shared one it used to rely on.

**The easy way:** menu option **2**, *Register the app in Entra ID*. It asks for your tenant name, creates the registration with the right permissions, and prints the client ID. Sign in as an administrator who can create app registrations and consent to permissions.

**Or run it yourself**, if you would rather see exactly what is happening:

```powershell
Register-PnPEntraIDAppForInteractiveLogin `
    -ApplicationName "PnP Migration" `
    -Tenant contoso.onmicrosoft.com `
    -SharePointDelegatePermissions "AllSites.FullControl" `
    -GraphDelegatePermissions "Group.Read.All","User.Read.All"
```

Replace `contoso.onmicrosoft.com` with your own tenant name. A browser opens for sign-in and consent. Add `-DeviceLogin` instead if there is no browser on this machine, or MFA makes the browser flow awkward.

**Keep this.** It prints a **client ID** — a long value like `66792cc7-5f4e-…`. Every SharePoint task asks for it. Write it down.

Consent can take a minute or two to apply. If the first SharePoint task fails on permissions, wait and try again.

> **Why the permissions are spelled out.** Left off, the command grants more than these scripts need, including write access to groups and users. The four scopes above are the actual requirement: full control of SharePoint sites, and read-only on groups and users so the owner report can resolve names.

### Step 5 — Optional: a certificate for every site

Skip this unless you need it.

Signing in as yourself only reaches sites you personally administer. Being SharePoint Administrator is *not* the same as being a site collection administrator, so on most team sites some reads will be refused.

If you need full coverage — and you will for demoting members across sites you don't own — create a certificate and let the scripts sign in as the application instead:

```powershell
cd Setup
./New-AppOnlyCertificate.ps1 -CommonName "PnP Migration" -ValidYears 1 `
    -OutPath ~/certs -CertificatePassword (Read-Host -AsSecureString)
```

Then in the Entra portal, open **App registrations → your app → Certificates & secrets → Certificates → Upload certificate** and upload the `.cer` file. Not the `.pfx` — that holds the private key and never leaves your machine.

The app also needs the application permission `Sites.FullControl.All`, granted and admin-consented in the portal.

> **Treat this seriously.** `Sites.FullControl.All` is unrestricted write access to every site in the tenant, held by a file on your disk. Password-protect the `.pfx`, keep it out of shared folders, and delete both it and the app registration when the migration is finished.

---

## Part 2 — Moving the guests

Do these in order. Each step depends on the one before it.

> **Before you start.** The groups must already exist in the new tenant, **with exactly the same display names**. That's what the import matches on — group IDs are different in every tenant and can't be used. Migrate your groups first.

### Step 1 — Export from the old tenant · [Read-only]

Menu option **4**. Sign in to the tenant you're moving *away from*. It reads every guest and the groups they belong to.

**What you get.** Two spreadsheets. `TenantA_GuestPermissions.csv` is the one the import reads. `…_Unsupported.csv` lists memberships that can't be recreated automatically — deal with those at step 6.

> **Fewer guests than expected?** Guests who are blocked from signing in, and guests who never accepted their original invitation, are left out by default. The wizard asks about both — say yes if you want them.

### Step 2 — Check the group names · [Read-only]

Open the export and look at the `GroupDisplayName` column. Every one of those names needs to exist in the new tenant, spelled identically.

This is the single most common cause of a half-finished migration.

### Step 3 — Rehearse the import · [Read-only]

Menu option **5**, signing in to the *new* tenant. The wizard previews automatically before doing anything. Nothing is created yet — but every group name is checked against the new tenant, which is the point.

**What to look for.** Open the log file it writes. Any row with status `GroupNotFound` is a group missing from the new tenant. Fix those before going further, or those memberships are silently skipped.

### Step 4 — Pilot with a handful · [Changes your tenant]

The wizard offers to start with a few guests. Take it — five is plenty. Say **no** to the invitation email for now.

**Then check by hand.** In the Entra admin centre, find one of those guests. Confirm they exist and are in the right groups. It's much cheaper to find a problem now than after two hundred invitations.

### Step 5 — Run it for everyone · [Changes your tenant]

Same task, no limit this time. The pilot guests are already there, so they're skipped rather than duplicated.

> **The invitation email is your decision.** By default guests are created silently — they appear in the directory and can be added to groups, but receive nothing. Say yes to the email only when you're ready for people to be told.
>
> Only guests *created on that run* are emailed. If you added guests by hand and now need to tell them, use menu option **6** instead. That sends a fresh invitation without creating a second account and without changing anyone's group memberships — worth knowing, because it is the obvious worry.

### Step 6 — Tell the guests · [Sends email]

If you created the guests silently — which is the sensible default — nothing has reached them yet. When you are ready, menu option **6**, *Email guests you have already created*.

It sends the invitation email to everyone in the spreadsheet and **changes no permissions at all**. Group membership is not touched, so anyone given extra groups since they were migrated keeps them, and no second account is created: the existing guest is matched on their email address.

Everyone listed in the file is emailed, so trim the spreadsheet first if some of them should not be told yet.

#### If nobody receives anything

Two things account for almost every case.

**1. The confirmation was not answered.** Every step that changes something runs twice: a rehearsal that sends nothing, then the real thing. In between you are asked *Go ahead and make these changes?* and the answer defaults to **no**, so pressing Enter sends nothing. You have to type `y`.

**2. The guest has already accepted.** Microsoft only emails an invitation to a guest who has not accepted one yet. Ask for another one for a guest who has already signed in, and the request succeeds quietly and no email is ever sent. Nothing looks wrong.

Step 6 asks *Re-invite guests who have already accepted*. Answering yes resets those guests so the email is delivered. They keep their account, their groups and their app access, but the next time they open a resource they are asked to accept the invitation again — and must do so with the address in the spreadsheet. Guests who have never accepted are emailed either way, so answering no is safe: the log tells you afterwards whether anyone was skipped.

**Read the log to tell which happened.** Open the results file and look at the `Status` column:

| Status | Meaning |
|---|---|
| `InvitationResent` | Microsoft accepted it and sent the email |
| `NoEmailSent` | The guest had already accepted. Run step 6 again and answer yes to re-inviting them |
| `WhatIf` | The rehearsal only. Nothing was sent — run it again and answer `y` when asked |
| `Failed` | Microsoft rejected it. The `Detail` column says why |

If the log says `InvitationResent` and the guest still has nothing, the email left Microsoft and the problem is at their end. It arrives from Microsoft Invitations, so junk mail, quarantine and the recipient's own tenant filtering are the places to look.

### Step 7 — Handle what's left over · [Changes your tenant]

Open `…_Unsupported.csv`. These need doing by hand, in Exchange Online PowerShell or the admin centre:

| Group type | Why it's here | What to do |
|---|---|---|
| Distribution list | Graph cannot add members to these | `Add-DistributionGroupMember` in Exchange Online |
| Mail-enabled security group | Same | Same |
| Dynamic group | Membership comes from a rule, not a member list | Recreate the rule in the new tenant |

---

## Part 3 — SharePoint access

Independent of the guest migration — run these whenever you need them.

### Find out who owns what · [Read-only]

Menu option **7**. Choose "every site in the tenant" and give it your admin URL — your SharePoint address with `-admin` inserted, so `contoso.sharepoint.com` becomes `contoso-admin.sharepoint.com`. Type the ordinary address by mistake and it is corrected for you, with a line saying so.

> **"Every site in the tenant" needs the SharePoint Administrator or Global Administrator role.** Only those roles may open the admin site, however much access you have to individual sites. Without one, the admin site answers with a sign-in page instead of data and the run stops with an explanation. You do not have to solve that to get a report — choose one site, or a spreadsheet of sites, instead. Neither goes near the admin site.

"Owner" means several different things in SharePoint, and the report lists all of them, so one person can appear more than once for the same site. That's intentional — see Reading the reports.

> **Watch for orphaned sites.** A row marked `OrphanedGroup` means the site is connected to a Microsoft 365 group that no longer exists. Those sites have no owners at all and no group to inherit them from. Assign someone before migrating anything.

### Find out who can get in · [Read-only]

Menu option **8**. The companion to the owners report: that one answers *who runs this site*, this one answers *who can get into it*.

Access reaches a site by more routes than most people expect, and all of them are listed, each row tagged with which route it came from:

| Route | What it means |
|---|---|
| `MembersGroup` | The site's Members group |
| `VisitorsGroup` | The Visitors group — read-only, but still access |
| `SharePointGroup` | Any other group on the site. Real sites accumulate these and they get forgotten |
| `DirectPermission` | Someone given permission on the site itself rather than through a group |
| `Microsoft365GroupMember` | Members of the connected Microsoft 365 group, on a Teams site. These people have access even when the Members group looks empty |

Somebody who has access two ways appears twice. That is deliberate — it is exactly what you have to unpick to take their access away.

The `Roles` column shows the permission level behind each row, read from the site's own settings, so a custom group with Full Control does not read as ordinary membership.

> **Before a migration, answer yes to "Guests only".** It reduces the report to external people and how each of them got in — usually the only version anyone reads.

Owners are left to the previous option, since that one also covers site collection administrators. Answer yes to *Include site owners as well* if you want one file covering everybody.

> **"Everyone except external users" is a finding, not noise.** It and the other tenant-wide claims are filtered out by default so they don't bury the real names. Worth running once with them included: a site carrying that claim is open far wider than its group membership suggests.

### Make members read-only · [Changes your tenant]

Menu option **9**. Demotes people from edit to read-only on a site, in both places access comes from:

- **the site's Members group** — they move to Visitors
- **permissions given to them directly on the site** — reduced to read-only

Both matter. Moving someone out of Members changes nothing if they were also given Edit directly, which is easily done and easily forgotten. The wizard offers to leave direct permissions alone if you really want group membership only.

People are added to Visitors *before* being removed from Members, so nobody is ever left with no access.

You can do one site, a whole list, or every site in the tenant. Choosing the spreadsheet option asks for a `.csv` file with a single column headed `SiteUrl`, one site per row:

```
SiteUrl
https://contoso.sharepoint.com/sites/Marketing
https://contoso.sharepoint.com/sites/Projects
```

Other columns are ignored, so a site list exported from elsewhere usually works as-is. The wizard reads the file straight away, shows you how many sites it found, and asks you to confirm before going any further.

> **Choosing "every site in the tenant"** changes permissions across the whole tenant, so it asks you to confirm that specifically before anything else. Personal OneDrive sites are left out, and sites with no Visitors group are skipped rather than half-changed. You still get the full preview before anything is applied.
>
> Signing in as yourself only reaches sites you administer, so a tenant-wide run will fail on the rest. Use the certificate from Part 1 if you need to cover everything. Three questions decide the scope:

- **Who?** Guests only (the default), your own staff only, or both. Whichever you choose applies to group membership and direct permissions alike.
- **Move the Team entry?** On a Teams-connected site this is the important one — see below.
- **Anyone to leave alone?** Give the email addresses of service or break-glass accounts that must keep editing.

> **Teams sites work differently.** On a site connected to a Team, the Members list contains the Team itself — shown as *"&lt;Site&gt; Members"* — alongside any individuals. Moving that single entry makes the whole team read-only on that site in one step, including people who join later.
>
> It does **not** affect Teams chat, the group mailbox or the calendar. Microsoft 365 group membership is never modified, deliberately — removing someone there would strip all three.

---

### List all users and their group access · [Read-only]

Menu option **10**. A spreadsheet of every user in the tenant — staff and guests alike — with the groups each one belongs to.

This one shows **effective** access: a group someone reaches only through another group is included. That makes it right for a permissions review, and different on purpose from the guest export at step 1 of Part 2, which records direct memberships only.

Useful to run against both tenants — before the migration and after — as a record of what access looked like.

---

## Part 4 — Reading the reports

Everything writes a CSV. These are the values worth understanding.

### Guest import log

| Status | Meaning | Action |
|---|---|---|
| `Created` | New guest invited | None |
| `AlreadyExists` | Guest was already there and was reused | None — this is the re-run behaviour working |
| `Added` / `AlreadyMember` | Group membership created, or already present | None |
| `GroupNotFound` | No group of that name in the new tenant | **Create the group, then re-run** |
| `Skipped` | Can't be done via Graph, or the name is ambiguous | Check the Detail column |
| `Failed` | Graph rejected it | Read the Detail column |
| `WhatIf` | Rehearsal — nothing was changed | None |

### All users report

| Column | Meaning |
|---|---|
| `UserType` | `Member` (your staff) or `Guest` |
| `AccountStatus` | `Enabled` or `Disabled` |
| `GroupName` | Group display name, or `None` if they belong to none |
| `GroupType` | Microsoft 365 group, security group, or distribution group |

One row per user per group, so someone in five groups appears five times.

### Site owners report

| OwnerSource | What it actually is |
|---|---|
| `SiteCollectionAdmin` | Full control of the site collection. The real administrators — and the ones most often overlooked. |
| `OwnersGroup` | Members of the site's Owners group in SharePoint. |
| `Microsoft365GroupOwner` | Owners of the connected group. On a Teams site these *are* the owners, even when the Owners group looks empty. |
| `TenantSiteOwner` | The primary owner recorded against the site collection. |
| `OrphanedGroup` | The connected group is gone. This site has no owners — assign one. |
| `Error` | You couldn't open that site, so its group-level owners are missing from the report. |

Guest owners are flagged in the `IsGuest` column — worth reviewing before a migration. The `SHAREPOINT\system` account and the tenant-wide admin roles are filtered out by default, since they appear identically on nearly every site.

---

## Part 5 — When something goes wrong

Real errors, and what each one actually means. Search this section for the text you're seeing.

### "Input CSV is missing required column(s)…"

The columns are almost certainly there. The file didn't parse into columns at all, so every one looks missing. The error names the file and the columns it *did* find, which tells you which case you're in:

- **It's an Excel workbook wearing a `.csv` name.** Opening a CSV in Excel and using *Save As* writes a real `.xlsx` while leaving the extension alone. Re-export, or save as *CSV UTF-8 (Comma delimited)*.
- **The file has no line breaks.** Something flattened it in transit. It can't be recovered — re-export, and move it as a file rather than pasting its contents.
- **It's the wrong file.** The `Found:` line lists the real columns; compare them against what you meant to pass.

### "Attempted to perform an unauthorized operation."

Expected, not a fault. You're not a site collection administrator on that site — being SharePoint Administrator doesn't grant that automatically.

Only two of the owner sources need site access. Teams sites still report their real owners through Graph, and other sites still report their primary owner, so those rows aren't blank. For the full picture, either use the app-only certificate from Part 1, or add yourself as site collection administrator on the sites you care about and re-run.

### "The specified X509 certificate store does not exist."

You're on Linux or macOS and used a certificate thumbprint that isn't installed on this machine. Point at the file instead:

```powershell
-CertificatePath ~/certs/PnPMigration.pfx `
-CertificatePassword (Read-Host -AsSecureString)
```

That works identically on Windows too, which makes it the better habit.

### You're asked to sign in at every single site

Your PnP module is too old to cache the sign-in. Update it:

```powershell
Update-Module PnP.PowerShell
```

The scripts warn about this at startup when they detect it. **Check and install what is needed** — the first menu option — offers the update for you.

### "A parameter cannot be found that matches parameter name…"

Your copy of the scripts is older than the command you're running. Ask whoever supplied the toolkit for the current version, and replace the whole folder rather than individual files.

### "Could not find … Common/…"

You're running a script that's been separated from the rest of the toolkit. The scripts share code in the `Common` folder. Put the script back in the complete folder, or get a fresh copy of the whole thing.

### Nothing happens, or it asks for "SiteUrl[0]"

You ran a script directly without telling it which sites to work on. Use the menu instead — it asks the questions in the right order:

```powershell
./Start-MigrationToolkit.ps1
```

---

## Part 6 — Afterwards

- **Keep the CSVs.** They're your record of what the permissions were and what was done — worth archiving somewhere durable.
- **Delete the app registration** once the migration is done, particularly if you granted `Sites.FullControl.All`.
- **Delete the `.pfx` file.** It authenticates as the application. It has no other purpose once the app is gone.
- **Remove any temporary site collection admin rights** you granted yourself along the way.

---

**Running the scripts directly.** The menu is a convenience, not a requirement. If
you would rather call the scripts with parameters — to script a run, schedule one,
or use an option the menu doesn't expose — see the per-script reference in
`docs/scripts/`.

---

*Every task writes a CSV, and every task that changes your tenant rehearses first. If you're ever unsure what something will do, run it and read the preview — that's what it's for.*
