<#
.SYNOPSIS
    Shared PnP.PowerShell connection handling.

.DESCRIPTION
    Dot-sourced by the scripts in SharePoint/. Not run directly.

    Supports both sign-in modes, chosen by which parameters are supplied.

    INTERACTIVE (default)
      You sign in as yourself. Access is your access: PnP asks Entra for
      delegated permissions, and the effective rights are the app's permissions
      intersected with your own. Being SharePoint Administrator does NOT make you
      a site collection administrator, so reads inside sites you do not own are
      refused. Fine for a handful of sites you own.

      Repeated sign-in prompts are avoided two ways:

        1. -PersistLogin caches the delegated token in
           %LOCALAPPDATA%\.m365pnppowershell (or $HOME/.m365pnppowershell),
           encrypted with DPAPI on Windows and the Keychain elsewhere. Later
           connections reuse it, across sessions and reboots. The switch is only
           present in newer PnP.PowerShell releases, so it is detected, not
           assumed.

        2. Not calling Disconnect-PnPOnline between sites. Disconnecting drops the
           current context, so the next Connect starts from nothing and prompts
           again. Connecting to a new URL replaces the previous connection by
           itself.

      To forget a cached token: Disconnect-PnPOnline -ClearPersistedLogin.

    APP-ONLY (a certificate thumbprint or path is supplied)
      The script authenticates as the application, not as you, using application
      permissions granted to the app registration. With SharePoint
      Sites.FullControl.All the app reaches every site in the tenant regardless of
      who is a site collection administrator, and nothing prompts. This is the
      mode to use for a whole-tenant sweep.
#>

# Parameter-name -> supported, filled lazily by Test-PnPParameter.
$script:PnPParameterCache = @{}

function Test-PnPParameter {
    <#  True when the installed Connect-PnPOnline exposes the named parameter.
        Capabilities differ between PnP.PowerShell releases, so they are probed
        rather than assumed. #>
    param([Parameter(Mandatory)][string]$Name)

    if (-not $script:PnPParameterCache.ContainsKey($Name)) {

        $supported = $false

        try {
            $command = Get-Command -Name Connect-PnPOnline -ErrorAction Stop

            $supported = $command.Parameters.ContainsKey($Name)
        }
        catch {
            $supported = $false
        }

        $script:PnPParameterCache[$Name] = $supported
    }

    return $script:PnPParameterCache[$Name]
}

function Test-PnPPersistLoginSupport {
    <#  True when the installed PnP.PowerShell exposes -PersistLogin. #>

    return (Test-PnPParameter -Name 'PersistLogin')
}

function Test-OnWindows {
    <#  $IsWindows only exists in PowerShell 6+. Windows PowerShell 5.1 runs
        nowhere else, so treat its absence as Windows. #>

    if ($null -ne $IsWindows) { return [bool]$IsWindows }

    return $true
}

function Resolve-CertificateByThumbprint {
    <#  Finds a certificate by thumbprint using .NET directly, which works on every
        platform .NET supports - on Linux and macOS the CurrentUser store is backed
        by ~/.dotnet/corefx/cryptography/x509stores. Returns $null when not found.

        Thumbprints are compared with punctuation and spaces stripped, so a value
        copied out of the Windows certificate dialog (which is space-separated and
        may carry a leading invisible character) still matches. #>
    param([Parameter(Mandatory)][string]$Thumbprint)

    $wanted = ($Thumbprint -replace '[^0-9A-Fa-f]', '')

    if (-not $wanted) { return $null }

    foreach ($location in @('CurrentUser', 'LocalMachine')) {

        $store = $null

        try {
            $store = [System.Security.Cryptography.X509Certificates.X509Store]::new('My', $location)
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)

            foreach ($candidate in $store.Certificates) {
                if ($candidate.Thumbprint -and $candidate.Thumbprint -eq $wanted) { return $candidate }
            }
        }
        catch {
            # A store that does not exist on this platform is not an error here.
            continue
        }
        finally {
            if ($store) { $store.Dispose() }
        }
    }

    return $null
}

function New-ScriptAuthContext {
    <#  Validates the sign-in parameters once, up front, and returns the context
        Connect-ScriptSite uses for every site. Catching a bad combination here
        beats discovering it partway through a tenant sweep. #>
    param(
        [string]$ClientId,
        [string]$Tenant,
        [string]$Thumbprint,
        [string]$CertificatePath,
        [System.Security.SecureString]$CertificatePassword,
        [switch]$NoPersistedLogin
    )

    $appOnly     = [bool]($Thumbprint -or $CertificatePath)
    $certificate = $null

    if ($appOnly) {

        $missing = @()

        if (-not $ClientId) { $missing += '-ClientId' }
        if (-not $Tenant)   { $missing += '-Tenant' }

        if ($missing.Count -gt 0) {
            throw "App-only sign-in also needs $($missing -join ' and '). Example: -ClientId <app id> -Tenant contoso.onmicrosoft.com -Thumbprint <cert thumbprint>"
        }

        if ($Thumbprint -and $CertificatePath) {
            throw 'Supply either -Thumbprint or -CertificatePath, not both.'
        }

        # PnP resolves -Thumbprint through the Windows certificate store. On Linux
        # and macOS that lookup fails ("The specified X509 certificate store does
        # not exist"), so resolve the certificate here and hand PnP the object.
        if ($Thumbprint -and -not (Test-OnWindows)) {

            $resolved = Resolve-CertificateByThumbprint -Thumbprint $Thumbprint

            if (-not $resolved) {
                $message = [System.Text.StringBuilder]::new()

                [void]$message.AppendLine("No certificate with thumbprint '$Thumbprint' is available on this machine.")
                [void]$message.AppendLine('')
                [void]$message.AppendLine('  On Linux and macOS there is no Windows certificate store to look it up in.')
                [void]$message.AppendLine('  Point at the .pfx that Register-PnPEntraIDApp wrote to its -OutPath instead:')
                [void]$message.AppendLine('      -CertificatePath /path/to/yourapp.pfx')
                [void]$message.AppendLine('')
                [void]$message.AppendLine('  Add -CertificatePassword (Read-Host -AsSecureString) if the .pfx is protected.')
                [void]$message.AppendLine('  -CertificatePath works on Windows too, so it is the portable choice for a')
                [void]$message.AppendLine('  script that has to run on both.')

                throw $message.ToString()
            }

            if (-not (Test-PnPParameter -Name 'Certificate')) {
                throw "This version of PnP.PowerShell cannot accept a certificate object, which is how -Thumbprint is supported on Linux and macOS. Use -CertificatePath with the .pfx instead, or run 'Update-Module PnP.PowerShell'."
            }

            $certificate = $resolved
        }

        if ($CertificatePath -and -not (Test-Path -Path $CertificatePath)) {
            throw "Certificate file not found: $CertificatePath"
        }
    }

    [PSCustomObject]@{
        AppOnly             = $appOnly
        Certificate         = $certificate
        ClientId            = $ClientId
        Tenant              = $Tenant
        Thumbprint          = $Thumbprint
        CertificatePath     = $CertificatePath
        CertificatePassword = $CertificatePassword
        NoPersistedLogin    = [bool]$NoPersistedLogin
    }
}

function Connect-ScriptSite {
    <#  Connects to one site collection using the context from
        New-ScriptAuthContext. #>
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)]$Auth
    )

    $params = @{
        Url         = $Url
        ErrorAction = 'Stop'
    }

    if ($Auth.ClientId) { $params['ClientId'] = $Auth.ClientId }

    if ($Auth.AppOnly) {

        $params['Tenant'] = $Auth.Tenant

        if ($Auth.Certificate) {
            # Resolved by us because this platform has no Windows certificate store.
            $params['Certificate'] = $Auth.Certificate
        }
        elseif ($Auth.Thumbprint) {
            $params['Thumbprint'] = $Auth.Thumbprint
        }
        else {
            $params['CertificatePath'] = $Auth.CertificatePath

            if ($Auth.CertificatePassword) {
                $params['CertificatePassword'] = $Auth.CertificatePassword
            }
        }
    }
    else {

        $params['Interactive'] = $true

        if (-not $Auth.NoPersistedLogin -and (Test-PnPPersistLoginSupport)) {
            $params['PersistLogin'] = $true
        }
    }

    Connect-PnPOnline @params
}

function Write-PnPLoginAdvice {
    <#  Says once, up front, how this run will authenticate and what that means
        for coverage. #>
    param([Parameter(Mandatory)]$Auth)

    if ($Auth.AppOnly) {
        Write-Host '  App-only sign-in. Coverage depends on the application permissions granted to the app, not on your own site access.' -ForegroundColor DarkGray
        return
    }

    Write-Host '  Interactive sign-in: sites you are not a site collection administrator of will refuse some reads.' -ForegroundColor DarkGray
    Write-Host '  For full coverage across a tenant, use app-only - see the README.' -ForegroundColor DarkGray

    if ($Auth.NoPersistedLogin) {
        Write-Host '  -NoPersistedLogin set: you will be asked to sign in for each site.' -ForegroundColor Yellow
        return
    }

    if (Test-PnPPersistLoginSupport) {
        Write-Host '  Sign-in is cached after the first site (-PersistLogin).' -ForegroundColor DarkGray
        return
    }

    Write-Warning 'This version of PnP.PowerShell has no -PersistLogin, so every site will ask you to sign in. Update the module to avoid that:'
    Write-Warning '    Update-Module PnP.PowerShell'
}
