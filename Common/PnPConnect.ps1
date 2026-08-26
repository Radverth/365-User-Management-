<#
.SYNOPSIS
    Shared PnP.PowerShell connection handling.

.DESCRIPTION
    Dot-sourced by the scripts in SharePoint/. Not run directly.

    Scripts here connect to one site collection after another, and a naive loop
    asks the user to sign in at every single one. Two things prevent that:

      1. -PersistLogin caches the delegated token in
         %LOCALAPPDATA%\.m365pnppowershell (or $HOME/.m365pnppowershell), encrypted
         with DPAPI on Windows and the Keychain elsewhere. Later connections reuse
         it, including in a fresh PowerShell session or after a reboot. The switch
         is only present in newer PnP.PowerShell releases, so it is detected rather
         than assumed.

      2. Not calling Disconnect-PnPOnline between sites. Disconnecting drops the
         current context, so the next Connect starts from nothing and prompts
         again. Connecting to a new URL replaces the previous connection on its
         own; disconnect once at the end of the run, if at all.

    To forget a cached token, run Disconnect-PnPOnline -ClearPersistedLogin.
#>

# Cached result of the capability probe below; $null until first checked.
$script:PnPPersistLoginSupported = $null

function Test-PnPPersistLoginSupport {
    <#  True when the installed PnP.PowerShell exposes -PersistLogin. #>

    if ($null -ne $script:PnPPersistLoginSupported) { return $script:PnPPersistLoginSupported }

    try {
        $command = Get-Command -Name Connect-PnPOnline -ErrorAction Stop

        $script:PnPPersistLoginSupported = $command.Parameters.ContainsKey('PersistLogin')
    }
    catch {
        $script:PnPPersistLoginSupported = $false
    }

    return $script:PnPPersistLoginSupported
}

function Connect-ScriptSite {
    <#  Interactive connection to one site collection, reusing a cached token where
        the installed PnP.PowerShell can. #>
    param(
        [Parameter(Mandatory)][string]$Url,

        [string]$ClientId,

        # Force a fresh sign-in instead of reusing the persisted token.
        [switch]$NoPersistedLogin
    )

    $params = @{
        Url         = $Url
        Interactive = $true
        ErrorAction = 'Stop'
    }

    if ($ClientId) { $params['ClientId'] = $ClientId }

    if (-not $NoPersistedLogin -and (Test-PnPPersistLoginSupport)) {
        $params['PersistLogin'] = $true
    }

    Connect-PnPOnline @params
}

function Write-PnPLoginAdvice {
    <#  Says once, up front, whether repeated sign-in prompts are expected. #>
    param([switch]$NoPersistedLogin)

    if ($NoPersistedLogin) {
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
