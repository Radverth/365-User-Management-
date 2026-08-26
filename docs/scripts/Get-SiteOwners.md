# Get-SiteOwners.ps1

Reports the owners of SharePoint sites to CSV, covering every meaning SharePoint gives the word.

**Changes nothing.** Read-only.

```
SharePoint/Get-SiteOwners.ps1
```

| | |
|---|---|
| **Modules** | `PnP.PowerShell` |
| **Loads** | `Common/InputCsv.ps1`, `Common/PnPConnect.ps1` |
| **Needs** | An Entra app registration (see [New-AppOnlyCertificate](New-AppOnlyCertificate.md) or the README) |
| **Supports `-WhatIf`** | No — it writes nothing to the tenant |

Running it with no site selection prints the three usage forms rather than prompting.

---

## Choosing sites — three mutually exclusive forms

| Parameter set | Parameters |
|---|---|
| One or more URLs | `-SiteUrl <string[]>` |
| From a spreadsheet | `-SitesCsvPath <file>` — needs a `SiteUrl` column |
| Whole tenant | `-AllSites -TenantAdminUrl <url>` (both required together) |

## Parameters

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-SiteUrl` | string[] | — | One or more site collection URLs |
| `-SitesCsvPath` | string | — | CSV with a `SiteUrl` column |
| `-AllSites` | switch | — | Every site in the tenant |
| `-TenantAdminUrl` | string | — | Required with `-AllSites`, e.g. `https://contoso-admin.sharepoint.com` |
| `-IncludeOneDrive` | switch | off | With `-AllSites`, also include personal OneDrive sites |
| `-ClientId` | string | — | Entra app client ID |
| `-Tenant` | string | — | Tenant name. Required for app-only sign-in |
| `-Thumbprint` | string | — | Certificate thumbprint. Supplying it switches to app-only |
| `-CertificatePath` | string | — | `.pfx` instead of a thumbprint. Also switches to app-only |
| `-CertificatePassword` | SecureString | — | Password for the `.pfx` |
| `-NoPersistedLogin` | switch | off | Sign in afresh at each site instead of reusing a cached token |
| `-IncludeSiteCollectionAdmins` | bool | `$true` | Set `:$false` to omit that source |
| `-IncludeOwnersGroup` | bool | `$true` | Set `:$false` to omit that source |
| `-IncludeMicrosoft365GroupOwners` | bool | `$true` | Set `:$false` to omit that source |
| `-IncludeSystemPrincipals` | switch | off | Keep `SHAREPOINT\system` and the tenant admin role claims |
| `-OutputPath` | string | `.\SharePoint_SiteOwners.csv` | The report |
| `-Delimiter` | string | auto | Field separator of `-SitesCsvPath` |

## Examples

```powershell
# Whole tenant, signing in as yourself
./Get-SiteOwners.ps1 -AllSites `
    -TenantAdminUrl https://contoso-admin.sharepoint.com -ClientId <app id>

# Whole tenant, app-only - reaches sites you do not administer
./Get-SiteOwners.ps1 -AllSites `
    -TenantAdminUrl https://contoso-admin.sharepoint.com `
    -ClientId <app id> -Tenant contoso.onmicrosoft.com `
    -CertificatePath ./PnPMigration.pfx `
    -CertificatePassword (Read-Host -AsSecureString)

# Two named sites, site-collection admins only
./Get-SiteOwners.ps1 -ClientId <app id> `
    -SiteUrl https://contoso.sharepoint.com/sites/A,https://contoso.sharepoint.com/sites/B `
    -IncludeOwnersGroup:$false -IncludeMicrosoft365GroupOwners:$false
```

---

## Output columns

`SiteUrl`, `SiteTitle`, `Template`, `IsGroupConnected`, `OwnerSource`, `OwnerName`, `OwnerLogin`, `OwnerEmail`, `PrincipalType`, `IsGuest`, `Note`

One row per owner per source, so the same person legitimately appears more than once for a site.

| `OwnerSource` | What it is |
|---|---|
| `SiteCollectionAdmin` | Full control of the site collection. The real administrators, and the ones most often overlooked |
| `OwnersGroup` | Members of the site's associated Owners SharePoint group |
| `Microsoft365GroupOwner` | Owners of the connected group. On a Teams site these *are* the owners, even when the Owners group looks empty |
| `TenantSiteOwner` | Primary owner from the tenant listing. `-AllSites` only |
| `OrphanedGroup` | The site is group-connected but the group no longer exists. No owners to inherit — assign one |
| `Error` | The site could not be opened, so its group-level owners are absent |
| `None` | No owners were found at all |

---

## Behaviour worth knowing

**Interactive sign-in only sees what you see.** Being SharePoint Administrator does not make you a site collection administrator, so `SiteCollectionAdmin` and `OwnersGroup` reads are refused on sites you do not own — logged as `Error` with *"Attempted to perform an unauthorized operation"*. The other sources still report, so those sites are not blank. App-only sign-in with `Sites.FullControl.All` removes the limitation.

**`-AllSites` keeps the tenant listing.** Title, template and primary owner come from the tenant enumeration, so a site that refuses to open still yields a `TenantSiteOwner` row rather than vanishing.

**Microsoft 365 group owners are read the long way round.** `Get-PnPMicrosoft365GroupOwner` returns only the object ID and leaves the name, UPN and mail blank ([pnp/powershell#5069](https://github.com/pnp/powershell/issues/5069)), so the script uses `Get-PnPMicrosoft365Group -IncludeOwners` instead. If a row still shows only a GUID, its `Note` says so — grant the app `User.Read.All`.

**Noise is removed by default and counted.** `SHAREPOINT\system` is never a person, and the tenant-wide Global Administrator / SharePoint Administrator role claims appear identically on most sites. `-IncludeSystemPrincipals` keeps them. Duplicate rows for the same owner on the same site are collapsed.

**Expect it to take a while.** `-AllSites` opens each site in turn. Seconds per site, not milliseconds.
