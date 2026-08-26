<#
.SYNOPSIS
    Creates a self-signed certificate for app-only sign-in, on any platform.

.DESCRIPTION
    Writes a .pfx (private key, used by the scripts) and a .cer (public key, uploaded
    to the Entra app registration), and optionally installs the certificate into your
    CurrentUser store so -Thumbprint works.

    Uses .NET's certificate API directly rather than New-SelfSignedCertificate
    (Windows only) or New-PnPAzureCertificate (which fails on some PowerShell 7.4
    builds - pnp/powershell#3838), so it behaves the same on Windows, Linux and
    macOS.

    Nothing here talks to Entra. Upload the .cer yourself:
      Entra portal -> App registrations -> your app -> Certificates & secrets
      -> Certificates -> Upload certificate

    That is also how you rotate an expiring certificate without touching the app's
    client ID: upload the new .cer alongside the old one, switch the scripts over,
    then delete the old entry.

.PARAMETER CommonName
    Subject common name. Cosmetic - it only identifies the certificate in listings.

.PARAMETER ValidYears
    Lifetime in years. Defaults to 1. Keep it short for a migration certificate.

.PARAMETER OutPath
    Directory to write into. Defaults to the current directory.

.PARAMETER BaseName
    File name stem for the .pfx and .cer. Defaults to the common name with spaces
    removed.

.PARAMETER CertificatePassword
    Protects the .pfx. Strongly recommended - the .pfx holds the private key.
    Prompt for it with (Read-Host -AsSecureString).

.PARAMETER Install
    Also add the certificate to your CurrentUser certificate store, so the scripts
    can find it by -Thumbprint. Without this, use -CertificatePath.

.PARAMETER Force
    Overwrite existing files with the same names.

.EXAMPLE
    .\New-AppOnlyCertificate.ps1 -CertificatePassword (Read-Host -AsSecureString)

.EXAMPLE
    .\New-AppOnlyCertificate.ps1 -CommonName "PnP Migration" -ValidYears 1 -Install `
        -OutPath ~/certs -CertificatePassword (Read-Host -AsSecureString)
#>

[CmdletBinding()]
param(
    [string]$CommonName = 'PnP Migration App-Only',

    [ValidateRange(1, 30)]
    [int]$ValidYears = 1,

    [string]$OutPath = '.',

    [string]$BaseName,

    [System.Security.SecureString]$CertificatePassword,

    [switch]$Install,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

if (-not $BaseName) { $BaseName = ($CommonName -replace '[^\w\-]', '') }

if (-not (Test-Path -Path $OutPath)) {
    New-Item -ItemType Directory -Path $OutPath -Force | Out-Null
}

$resolvedOut = (Resolve-Path -Path $OutPath).Path
$pfxPath     = Join-Path -Path $resolvedOut -ChildPath "$BaseName.pfx"
$cerPath     = Join-Path -Path $resolvedOut -ChildPath "$BaseName.cer"

foreach ($existing in @($pfxPath, $cerPath)) {
    if ((Test-Path -Path $existing) -and -not $Force) {
        throw "$existing already exists. Use -Force to overwrite, or choose another -BaseName."
    }
}

if (-not $CertificatePassword) {
    Write-Warning 'No -CertificatePassword given, so the .pfx will be unprotected. It contains the private key - anyone holding it can authenticate as this application.'
}

Write-Host "Generating a $ValidYears-year certificate for '$CommonName'..." -ForegroundColor Cyan

$rsa = [System.Security.Cryptography.RSA]::Create(2048)

try {
    $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
        "CN=$CommonName",
        $rsa,
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
    )

    # Not a CA, and usable for the signing Entra expects of a client credential.
    $request.CertificateExtensions.Add(
        [System.Security.Cryptography.X509Certificates.X509BasicConstraintsExtension]::new($false, $false, 0, $true))

    $request.CertificateExtensions.Add(
        [System.Security.Cryptography.X509Certificates.X509KeyUsageExtension]::new(
            [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::DigitalSignature -bor
            [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::KeyEncipherment,
            $true))

    # Backdated slightly so clock skew between here and Entra cannot make a
    # freshly-issued certificate look not-yet-valid.
    $notBefore = [System.DateTimeOffset]::UtcNow.AddMinutes(-5)
    $notAfter  = $notBefore.AddYears($ValidYears)

    $certificate = $request.CreateSelfSigned($notBefore, $notAfter)
}
finally {
    $rsa.Dispose()
}

# SecureString -> plain text only for the moment of export, then wiped.
$plainPassword = ''

if ($CertificatePassword) {
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($CertificatePassword)

    try   { $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

try {

    $pfxBytes = $certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, $plainPassword)
    [System.IO.File]::WriteAllBytes($pfxPath, $pfxBytes)

    $cerBytes = $certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
    [System.IO.File]::WriteAllBytes($cerPath, $cerBytes)

    if ($Install) {

        $store = $null

        try {
            # Re-import from the pfx so the private key is persisted with the copy
            # that goes into the store; the in-memory object's key may not be.
            $flags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable -bor
                     [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::PersistKeySet

            $installable = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($pfxBytes, $plainPassword, $flags)

            $store = [System.Security.Cryptography.X509Certificates.X509Store]::new('My', 'CurrentUser')
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
            $store.Add($installable)

            Write-Host '  Installed into the CurrentUser certificate store.' -ForegroundColor Green
        }
        catch {
            Write-Warning "Could not install into the certificate store: $($_.Exception.Message)"
            Write-Warning 'Use -CertificatePath with the .pfx instead of -Thumbprint.'
        }
        finally {
            if ($store) { $store.Dispose() }
        }
    }
}
finally {
    # Held until here because installing into the store needs it to re-open the pfx.
    $plainPassword = $null
}

Write-Host ''
Write-Host '--- Certificate created ---' -ForegroundColor Green
Write-Host "  Thumbprint : $($certificate.Thumbprint)"
Write-Host "  Expires    : $($certificate.NotAfter.ToString('yyyy-MM-dd'))"
Write-Host "  Private key: $pfxPath" -ForegroundColor Yellow
Write-Host "  Public key : $cerPath"
Write-Host ''
Write-Host '--- Next ---' -ForegroundColor Cyan
Write-Host '  1. Upload the .cer to your app registration:'
Write-Host '       Entra portal -> App registrations -> your app -> Certificates & secrets'
Write-Host '       -> Certificates -> Upload certificate'
Write-Host ''
Write-Host '  2. Run the SharePoint scripts against it:'

if ($Install) {
    Write-Host "       -ClientId <app id> -Tenant <tenant>.onmicrosoft.com -Thumbprint $($certificate.Thumbprint)"
    Write-Host "     or, needing nothing installed:"
}

Write-Host "       -ClientId <app id> -Tenant <tenant>.onmicrosoft.com -CertificatePath $pfxPath"

if ($CertificatePassword) {
    Write-Host '       -CertificatePassword (Read-Host -AsSecureString)'
}

Write-Host ''
Write-Host "  Keep $pfxPath safe - it authenticates as the application. Delete it, and" -ForegroundColor Yellow
Write-Host '  the app registration, once the migration is done.' -ForegroundColor Yellow
