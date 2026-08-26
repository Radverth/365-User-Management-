#Requires -Modules PnP.PowerShell

<#
.SYNOPSIS
    Reports the owners of SharePoint sites to CSV.

.DESCRIPTION
    Read-only. Nothing is modified.

    "Owner" means three different things in SharePoint, and this script collects all
    of them so the report is complete:

      SiteCollectionAdmin  Full control over the whole site collection. The real
                           administrators, and the ones most often missed.
      OwnersGroup          Members of the site's associated Owners SharePoint group.
      Microsoft365Group    Owners of the connected Microsoft 365 group, for
                           group-connected (Teams) sites. These people are site
                           owners even when the Owners group looks empty.

    Each owner produces one row, tagged with which source it came from, so a person
    who is both a site collection admin and in the Owners group appears twice.

.PARAMETER SiteUrl
    One or more site collection URLs to report on.

.PARAMETER SitesCsvPath
    CSV containing a SiteUrl column, as an alternative to -SiteUrl.

.PARAMETER AllSites
    Report on every site in the tenant. Requires -TenantAdminUrl and SharePoint
    administrator rights.

    Each site is opened individually to read its groups, so allow time on a large
    tenant. Being SharePoint Administrator does not make you a site collection
    administrator everywhere; sites that refuse to open still produce a row with
    the primary owner from the tenant listing, marked with an Error row explaining
    that group-level owners are missing.

.PARAMETER TenantAdminUrl
    Your tenant admin URL, e.g. https://contoso-admin.sharepoint.com. Required with
    -AllSites.

.PARAMETER IncludeOneDrive
    With -AllSites, also include personal OneDrive sites. Excluded by default.

.PARAMETER ClientId
    Client ID of the Entra app registration used by PnP.PowerShell. See the README.

.PARAMETER IncludeSiteCollectionAdmins
    Include site collection administrators. Defaults to $true.

.PARAMETER IncludeOwnersGroup
    Include members of the associated Owners group. Defaults to $true.

.PARAMETER IncludeMicrosoft365GroupOwners
    Include owners of the connected Microsoft 365 group. Defaults to $true.

.PARAMETER OutputPath
    Destination CSV.

.PARAMETER Delimiter
    Field separator of -SitesCsvPath. Detected automatically when omitted.

.PARAMETER IncludeSystemPrincipals
    Keep principals that are not real owners: the SHAREPOINT\system account, and
    the tenant-wide Global Administrator / SharePoint Administrator role claims
    that appear as site collection administrators on most sites. Excluded by
    default, and the number removed is reported at the end.

.PARAMETER Tenant
    Tenant name, e.g. contoso.onmicrosoft.com. Required for app-only sign-in.

.PARAMETER Thumbprint
    Thumbprint of a certificate in your certificate store. Supplying it switches
    the script to app-only sign-in, where access comes from the app registration's
    application permissions rather than your own site access. With SharePoint
    Sites.FullControl.All this reaches every site in the tenant and never prompts.

.PARAMETER CertificatePath
    Path to a .pfx instead of -Thumbprint. Also switches to app-only sign-in.

.PARAMETER CertificatePassword
    Password for -CertificatePath, as a SecureString.

.PARAMETER NoPersistedLogin
    Sign in afresh for every site instead of reusing a cached token. Only needed
    when you must connect as different accounts, or to work around a stale token.

.EXAMPLE
    .\Get-SiteOwners.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Project -ClientId $id

.EXAMPLE
    .\Get-SiteOwners.ps1 -AllSites -TenantAdminUrl https://contoso-admin.sharepoint.com -ClientId $id

.EXAMPLE
    .\Get-SiteOwners.ps1 -SitesCsvPath .\sites.csv -ClientId $id -OutputPath .\owners.csv
#>

[CmdletBinding(DefaultParameterSetName = 'Urls')]
param(
    [Parameter(ParameterSetName = 'Urls')]
    [string[]]$SiteUrl,

    [Parameter(Mandatory, ParameterSetName = 'Csv')]
    [string]$SitesCsvPath,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch]$AllSites,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [string]$TenantAdminUrl,

    [Parameter(ParameterSetName = 'All')]
    [switch]$IncludeOneDrive,

    [string]$ClientId,

    [bool]$IncludeSiteCollectionAdmins = $true,

    [bool]$IncludeOwnersGroup = $true,

    [bool]$IncludeMicrosoft365GroupOwners = $true,

    [string]$OutputPath = ".\SharePoint_SiteOwners.csv",

    [string]$Delimiter,

    [string]$Tenant,

    [string]$Thumbprint,

    [string]$CertificatePath,

    [System.Security.SecureString]$CertificatePassword,

    [switch]$NoPersistedLogin,

    [switch]$IncludeSystemPrincipals
)

$ErrorActionPreference = 'Stop'

#region Helpers ---------------------------------------------------------------

function Initialize-OutputPath {
    param([Parameter(Mandatory)][string]$Path)

    $parent = Split-Path -Path $Path -Parent

    if ($parent -and -not (Test-Path -Path $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
}

# Shared CSV input handling - see Common/InputCsv.ps1
$commonPath = Join-Path -Path $PSScriptRoot -ChildPath '..\Common\InputCsv.ps1'

if (-not (Test-Path -Path $commonPath)) {
    throw "Could not find $commonPath. Run this script from inside a full clone of the repository - it depends on the shared helper in Common/."
}

. $commonPath

$pnpConnectPath = Join-Path -Path $PSScriptRoot -ChildPath '..\Common\PnPConnect.ps1'

if (-not (Test-Path -Path $pnpConnectPath)) {
    throw "Could not find $pnpConnectPath. Run this script from inside a full clone of the repository - it depends on the shared helpers in Common/."
}

. $pnpConnectPath

#endregion Helpers ------------------------------------------------------------

Initialize-OutputPath -Path $OutputPath

$auth = New-ScriptAuthContext -ClientId $ClientId -Tenant $Tenant -Thumbprint $Thumbprint `
                              -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword `
                              -NoPersistedLogin:$NoPersistedLogin

# Resolve the list of sites.
$sites = @()

# Url (lowercased, no trailing slash) -> the Get-PnPTenantSite record, when -AllSites
# was used. Empty in the other modes.
$tenantSiteIndex = @{}

switch ($PSCmdlet.ParameterSetName) {

    'Csv' {
        if (-not (Test-Path -Path $SitesCsvPath)) { throw "Sites CSV not found: $SitesCsvPath" }

        $csv = Import-InputCsv -Path $SitesCsvPath -Delimiter $Delimiter -RequiredColumns @('SiteUrl')

        if ($csv.Count -eq 0) { throw "Sites CSV is empty: $SitesCsvPath" }

        $sites = @($csv | Select-Object -ExpandProperty SiteUrl | Where-Object { $_ } | ForEach-Object { $_.Trim() })
    }

    'All' {
        Write-Host "Connecting to $TenantAdminUrl to enumerate sites..." -ForegroundColor Cyan

        Connect-ScriptSite -Url $TenantAdminUrl -Auth $auth

        $tenantSites = @(Get-PnPTenantSite -ErrorAction Stop)

        if (-not $IncludeOneDrive) {
            $tenantSites = @($tenantSites | Where-Object { $_.Template -notlike 'SPSPERS*' })
        }

        # The tenant listing already knows each site's title, template and primary
        # owner. Keep it: being SharePoint Administrator does not make you a site
        # collection admin everywhere, so some of these sites will refuse a direct
        # connection, and this is what lets those still produce an owner row.
        foreach ($tenantSite in $tenantSites) {
            if ($tenantSite.Url) {
                $tenantSiteIndex[$tenantSite.Url.TrimEnd('/').ToLowerInvariant()] = $tenantSite
            }
        }

        $sites = @($tenantSites | Select-Object -ExpandProperty Url)

        Write-Host "  $($sites.Count) site(s) found." -ForegroundColor Green
    }

    default {
        $sites = @($SiteUrl | Where-Object { $_ } | ForEach-Object { $_.Trim() })
    }
}

if ($sites.Count -eq 0) {

    $usage = [System.Text.StringBuilder]::new()

    [void]$usage.AppendLine('No sites to process. Choose one of the three modes:')
    [void]$usage.AppendLine('')
    [void]$usage.AppendLine('  Every site in the tenant:')
    [void]$usage.AppendLine('    .\Get-SiteOwners.ps1 -AllSites -TenantAdminUrl https://<tenant>-admin.sharepoint.com -ClientId <id>')
    [void]$usage.AppendLine('')
    [void]$usage.AppendLine('  Specific sites:')
    [void]$usage.AppendLine('    .\Get-SiteOwners.ps1 -SiteUrl https://<tenant>.sharepoint.com/sites/One,https://<tenant>.sharepoint.com/sites/Two -ClientId <id>')
    [void]$usage.AppendLine('')
    [void]$usage.AppendLine('  From a CSV with a SiteUrl column:')
    [void]$usage.AppendLine('    .\Get-SiteOwners.ps1 -SitesCsvPath .\sites.csv -ClientId <id>')

    throw $usage.ToString()
}

Write-Host "Reporting owners for $($sites.Count) site(s)..." -ForegroundColor Cyan
Write-PnPLoginAdvice -Auth $auth
Write-Host ''

$report = [System.Collections.Generic.List[PSCustomObject]]::new()

# Duplicate suppression and the per-site tally of real owner rows, both maintained
# by Add-OwnerRow.
$script:SeenOwnerRows           = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$script:SiteOwnerRowCount       = 0
$script:FilteredSystemPrincipals = 0

function Add-OwnerRow {
    param(
        [Parameter(Mandatory)][string]$Site,
        [string]$SiteTitle,
        [string]$Template,
        [string]$GroupConnected,
        [Parameter(Mandatory)][string]$Source,
        [string]$OwnerName,
        [string]$OwnerLogin,
        [string]$OwnerEmail,
        [string]$PrincipalType,
        [string]$IsGuest,
        [string]$Note = ''
    )

    # SHAREPOINT\system is never a person, and the tenant admin role claims are
    # identical on every site - both bury the real owners.
    if (-not $IncludeSystemPrincipals -and $OwnerLogin) {

        if ($OwnerLogin -eq 'SHAREPOINT\system' -or $OwnerLogin -like 'c:0t.c|tenant|*') {
            $script:FilteredSystemPrincipals++
            return
        }
    }

    # The same owner can surface more than once per site - a group listed by two
    # sources, or a repeated Graph page.
    $key = "$Site|$Source|$OwnerLogin|$OwnerName|$Note"

    if (-not $script:SeenOwnerRows.Add($key)) { return }

    if ($OwnerName -or $OwnerLogin) { $script:SiteOwnerRowCount++ }

    $report.Add([PSCustomObject]@{
        SiteUrl          = $Site
        SiteTitle        = $SiteTitle
        Template         = $Template
        IsGroupConnected = $GroupConnected
        OwnerSource      = $Source
        OwnerName        = $OwnerName
        OwnerLogin       = $OwnerLogin
        OwnerEmail       = $OwnerEmail
        PrincipalType    = $PrincipalType
        IsGuest          = $IsGuest
        Note             = $Note
    })
}

$counter = 0

foreach ($site in $sites) {

    $counter++

    Write-Progress -Activity 'Collecting site owners' `
                   -Status "$counter of $($sites.Count) - $site" `
                   -PercentComplete (($counter / $sites.Count) * 100)

    # Whatever the tenant listing already told us about this site.
    $tenantRecord = $tenantSiteIndex[$site.TrimEnd('/').ToLowerInvariant()]

    $siteTitle = if ($tenantRecord) { [string]$tenantRecord.Title } else { '' }
    $template  = if ($tenantRecord) { [string]$tenantRecord.Template } else { '' }
    $groupId   = $null

    if ($tenantRecord -and $tenantRecord.GroupId -and $tenantRecord.GroupId -ne [Guid]::Empty) {
        $groupId = $tenantRecord.GroupId
    }

    $script:SiteOwnerRowCount = 0

    # The primary owner recorded against the site collection itself. Available
    # without opening the site, so it survives a failed connection below.
    if ($tenantRecord -and $tenantRecord.Owner) {

        $ownerLogin = [string]$tenantRecord.Owner

        Add-OwnerRow -Site $site -SiteTitle $siteTitle -Template $template `
                     -GroupConnected ([bool]$groupId) -Source 'TenantSiteOwner' `
                     -OwnerName $tenantRecord.OwnerName -OwnerLogin $ownerLogin `
                     -OwnerEmail $tenantRecord.OwnerEmail -PrincipalType 'User' `
                     -IsGuest ($ownerLogin -imatch '(#ext#|urn:spo:guest)') `
                     -Note 'Primary owner from the tenant site listing'

    }

    try {
        Connect-ScriptSite -Url $site -Auth $auth
    }
    catch {
        # Being SharePoint Administrator does not grant access to every site, so
        # this is expected on some. Any tenant-listing owner above still stands.
        Write-Warning "Could not open $site : $($_.Exception.Message)"

        Add-OwnerRow -Site $site -SiteTitle $siteTitle -Template $template `
                     -GroupConnected ([bool]$groupId) -Source 'Error' `
                     -Note "Could not open the site, so group-level owners are missing: $($_.Exception.Message)"
        continue
    }

    try {
        $web = Get-PnPWeb -ErrorAction Stop

        # Only overwrite when the site actually gave us something. WebTemplate in
        # particular is often not loaded, and blanking the value from the tenant
        # listing loses information we already had.
        if ($web.Title)       { $siteTitle = [string]$web.Title }
        if ($web.WebTemplate) { $template  = [string]$web.WebTemplate }
    }
    catch {
        Write-Verbose "Could not read web properties for $site : $($_.Exception.Message)"
    }

    try {
        $pnpSite = Get-PnPSite -Includes GroupId -ErrorAction Stop

        if ($pnpSite.GroupId -and $pnpSite.GroupId -ne [Guid]::Empty) { $groupId = $pnpSite.GroupId }
    }
    catch {
        Write-Verbose "Could not read GroupId for $site : $($_.Exception.Message)"
    }

    $isGroupConnected = [bool]$groupId

    # --- site collection administrators -------------------------------------

    if ($IncludeSiteCollectionAdmins) {
        try {
            foreach ($admin in @(Get-PnPSiteCollectionAdmin -ErrorAction Stop)) {

                $login = [string]$admin.LoginName

                Add-OwnerRow -Site $site -SiteTitle $siteTitle -Template $template `
                             -GroupConnected $isGroupConnected -Source 'SiteCollectionAdmin' `
                             -OwnerName $admin.Title -OwnerLogin $login -OwnerEmail $admin.Email `
                             -PrincipalType $admin.PrincipalType `
                             -IsGuest ($login -imatch '(#ext#|urn:spo:guest)')

            }
        }
        catch {
            Write-Warning "  Site collection admins unavailable for $site : $($_.Exception.Message)"
            Add-OwnerRow -Site $site -SiteTitle $siteTitle -Template $template `
                         -GroupConnected $isGroupConnected -Source 'SiteCollectionAdmin' `
                         -Note "Unavailable: $($_.Exception.Message)"
        }
    }

    # --- associated Owners group --------------------------------------------

    if ($IncludeOwnersGroup) {
        try {
            $ownersGroup = Get-PnPGroup -AssociatedOwnerGroup -ErrorAction Stop

            if ($ownersGroup) {
                foreach ($owner in @(Get-PnPGroupMember -Identity $ownersGroup -ErrorAction Stop)) {

                    $login = [string]$owner.LoginName

                    Add-OwnerRow -Site $site -SiteTitle $siteTitle -Template $template `
                                 -GroupConnected $isGroupConnected -Source 'OwnersGroup' `
                                 -OwnerName $owner.Title -OwnerLogin $login -OwnerEmail $owner.Email `
                                 -PrincipalType $owner.PrincipalType `
                                 -IsGuest ($login -imatch '(#ext#|urn:spo:guest)') `
                                 -Note $ownersGroup.Title

                }
            }
        }
        catch {
            Write-Warning "  Owners group unavailable for $site : $($_.Exception.Message)"
            Add-OwnerRow -Site $site -SiteTitle $siteTitle -Template $template `
                         -GroupConnected $isGroupConnected -Source 'OwnersGroup' `
                         -Note "Unavailable: $($_.Exception.Message)"
        }
    }

    # --- Microsoft 365 group owners -----------------------------------------

    if ($IncludeMicrosoft365GroupOwners -and $groupId) {
        try {
            $groupOwners = @()

            # Get-PnPMicrosoft365GroupOwner populates only Id, leaving DisplayName,
            # UserPrincipalName and Mail empty (pnp/powershell#5069). Asking the
            # group to include its owners returns fully populated objects.
            try {
                $m365Group = Get-PnPMicrosoft365Group -Identity $groupId -IncludeOwners -ErrorAction Stop

                if ($m365Group -and $m365Group.Owners) { $groupOwners = @($m365Group.Owners) }
            }
            catch {
                # A missing group is a real finding (orphaned site) - let it reach
                # the handler below rather than falling back and reporting the
                # blank rows the buggy cmdlet returns.
                if ($_.Exception.Message -match 'does not exist|Not Found|\(404\)') { throw }

                $groupOwners = @(Get-PnPMicrosoft365GroupOwner -Identity $groupId -ErrorAction Stop)
            }

            if ($groupOwners.Count -eq 0) {
                $groupOwners = @(Get-PnPMicrosoft365GroupOwner -Identity $groupId -ErrorAction Stop)
            }

            foreach ($owner in $groupOwners) {

                $ownerName  = [string]$owner.DisplayName
                $ownerLogin = [string]$owner.UserPrincipalName
                $ownerMail  = [string]$owner.Mail

                if (-not $ownerMail)  { $ownerMail  = [string]$owner.Email }
                if (-not $ownerLogin) { $ownerLogin = $ownerMail }

                $note = "Microsoft 365 group $groupId"

                # Fall back to the directory object ID so the row still identifies
                # someone, rather than being a blank line in the report.
                if (-not $ownerName -and -not $ownerLogin) {

                    $ownerLogin = [string]$owner.Id

                    if ($ownerLogin) {
                        $note = "Microsoft 365 group $groupId; only the directory object ID was returned - grant the app User.Read.All to resolve names"
                    }
                }

                Add-OwnerRow -Site $site -SiteTitle $siteTitle -Template $template `
                             -GroupConnected $isGroupConnected -Source 'Microsoft365GroupOwner' `
                             -OwnerName $ownerName -OwnerLogin $ownerLogin -OwnerEmail $ownerMail `
                             -PrincipalType 'User' `
                             -IsGuest ($ownerLogin -imatch '#EXT#') `
                             -Note $note

            }
        }
        catch {
            $reason = $_.Exception.Message

            # A 404 here means the site outlived the group it was connected to.
            if ($reason -match 'does not exist|Not Found|\(404\)') {
                Write-Warning "  $site is connected to Microsoft 365 group $groupId, which no longer exists."

                Add-OwnerRow -Site $site -SiteTitle $siteTitle -Template $template `
                             -GroupConnected $isGroupConnected -Source 'OrphanedGroup' `
                             -Note "Connected Microsoft 365 group $groupId no longer exists. The site has no group to inherit owners from and needs a new owner."
            }
            else {
                Write-Warning "  Microsoft 365 group owners unavailable for $site : $reason"

                Add-OwnerRow -Site $site -SiteTitle $siteTitle -Template $template `
                             -GroupConnected $isGroupConnected -Source 'Microsoft365GroupOwner' `
                             -Note "Unavailable (Graph permission may be missing): $reason"
            }
        }
    }

    if ($script:SiteOwnerRowCount -eq 0) {
        Add-OwnerRow -Site $site -SiteTitle $siteTitle -Template $template `
                     -GroupConnected $isGroupConnected -Source 'None' `
                     -Note 'No owners were found for this site'
    }

    Write-Host ("[{0}/{1}] {2} - {3} owner(s)" -f $counter, $sites.Count, $site, $script:SiteOwnerRowCount) -ForegroundColor DarkGray

}

# One disconnect at the end. Doing it per site drops the token context and makes
# the next Connect prompt again. The persisted cache is left intact - clear it
# with Disconnect-PnPOnline -ClearPersistedLogin.
try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }

Write-Progress -Activity 'Collecting site owners' -Completed

$report |
    Sort-Object SiteUrl, OwnerSource, OwnerName |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host ''
Write-Host '--- Summary ---' -ForegroundColor Green
Write-Host "  Sites processed : $($sites.Count)"
Write-Host "  Owner rows      : $($report.Count)"

$report | Group-Object OwnerSource | Sort-Object Name | ForEach-Object {
    Write-Host ("    {0,-26} {1}" -f $_.Name, $_.Count)
}

if ($script:FilteredSystemPrincipals -gt 0) {
    Write-Host "  Filtered out    : $($script:FilteredSystemPrincipals) system / tenant-admin-role principal(s). Use -IncludeSystemPrincipals to keep them." -ForegroundColor DarkGray
}

$orphaned = @($report | Where-Object { $_.OwnerSource -eq 'OrphanedGroup' })

if ($orphaned.Count -gt 0) {
    Write-Host "  Orphaned sites  : $($orphaned.Count) connected to a Microsoft 365 group that no longer exists." -ForegroundColor Yellow
}

$guestOwners = @($report | Where-Object { $_.IsGuest -eq $true -or $_.IsGuest -eq 'True' })

if ($guestOwners.Count -gt 0) {
    Write-Host "  Guest owners    : $($guestOwners.Count)" -ForegroundColor Yellow
}

Write-Host ''
Write-Host "  Report : $((Resolve-Path -Path $OutputPath).Path)" -ForegroundColor Green
