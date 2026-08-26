# New-AppOnlyCertificate.ps1

Creates a self-signed certificate for app-only sign-in, on any platform.

**Writes local files only.** Nothing is sent to Entra — uploading the public key is a manual step.

```
Setup/New-AppOnlyCertificate.ps1
```

| | |
|---|---|
| **Modules** | None |
| **Loads** | Nothing — fully standalone, so it can be copied out of the folder |
| **Supports `-WhatIf`** | No |

Uses .NET's certificate API directly rather than `New-SelfSignedCertificate` (Windows only) or `New-PnPAzureCertificate` (fails on some PowerShell 7.4 builds — [pnp/powershell#3838](https://github.com/pnp/powershell/discussions/3838)), so it behaves the same on Windows, Linux and macOS.

---

## Parameters

| Parameter | Type | Default | Notes |
|---|---|---|---|
| `-CommonName` | string | `PnP Migration App-Only` | Subject name. Cosmetic — it identifies the certificate in listings |
| `-ValidYears` | int | `1` | Lifetime, 1–30. Keep it short for a migration certificate |
| `-OutPath` | string | `.` | Folder to write into. Created if missing |
| `-BaseName` | string | derived from `-CommonName` | Filename stem for the `.pfx` and `.cer` |
| `-CertificatePassword` | SecureString | — | Protects the `.pfx`. Strongly recommended |
| `-Install` | switch | off | Also add it to your CurrentUser certificate store, so `-Thumbprint` resolves |
| `-Force` | switch | off | Overwrite existing files of the same name |

## Examples

```powershell
# Recommended
./New-AppOnlyCertificate.ps1 -CertificatePassword (Read-Host -AsSecureString)

# Named, installed locally, in a chosen folder
./New-AppOnlyCertificate.ps1 -CommonName "PnP Migration" -ValidYears 1 `
    -OutPath ~/certs -Install -CertificatePassword (Read-Host -AsSecureString)

# Replace an expiring certificate
./New-AppOnlyCertificate.ps1 -CommonName "PnP Migration" -OutPath ~/certs -Force `
    -CertificatePassword (Read-Host -AsSecureString)
```

---

## What it produces

| File | Contains | Goes where |
|---|---|---|
| `<BaseName>.pfx` | Private key | Stays on your machine. Passed to the scripts as `-CertificatePath` |
| `<BaseName>.cer` | Public key | Uploaded to the app registration |

It prints the thumbprint, the expiry date, and the exact arguments to pass to the SharePoint scripts.

---

## Finishing the setup

The certificate is useless until Entra knows about it:

1. **Entra portal → App registrations → your app → Certificates & secrets → Certificates → Upload certificate**
2. Upload the **`.cer`**. Never the `.pfx`.
3. Under **API permissions**, add the SharePoint *application* permission `Sites.FullControl.All` and grant admin consent.

Then:

```powershell
-ClientId <app id> -Tenant contoso.onmicrosoft.com -CertificatePath ./yourapp.pfx `
-CertificatePassword (Read-Host -AsSecureString)
```

**To rotate without changing the app's client ID:** create a new certificate, upload the new `.cer` alongside the old one, switch the scripts over, then delete the old entry.

---

## Behaviour worth knowing

**`-Install` is verified, not assumed.** The store can accept a certificate and still not hand it back — a redirected or read-only store, or a different `$HOME` under `sudo`. The script reopens the store read-only and confirms the thumbprint is readable. If not, it warns and stops offering `-Thumbprint` in its closing instructions.

**`-Thumbprint` works on all three platforms,** but only if the certificate is installed on that machine. PnP resolves thumbprints through the Windows certificate store; on Linux and macOS the scripts resolve it through .NET instead. `-CertificatePath` needs nothing installed and is the portable choice.

**Existing files are never silently overwritten** — `-Force` is required.

**The certificate is backdated five minutes,** so clock skew between your machine and Entra cannot make a freshly issued certificate look not-yet-valid.

> **Handle the `.pfx` carefully.** With `Sites.FullControl.All` it grants unrestricted write access to every site in the tenant. Password-protect it, keep it out of shared locations, and delete it along with the app registration when the migration is done. The repository's `.gitignore` excludes `*.pfx` and `*.cer` so they cannot be committed by accident.
