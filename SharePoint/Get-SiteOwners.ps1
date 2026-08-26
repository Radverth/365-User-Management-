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

    [string]$Delimiter
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

function Connect-Site {
    param([Parameter(Mandatory)][string]$Url, [string]$AppId)

    $params = @{
        Url         = $Url
        Interactive = $true
        ErrorAction = 'Stop'
    }

    if ($AppId) { $params['ClientId'] = $AppId }

    Connect-PnPOnline @params
}

# Shared CSV input handling - see Common/InputCsv.ps1
$commonPath = Join-Path -Path $PSScriptRoot -ChildPath '..\Common\InputCsv.ps1'

if (-not (Test-Path -Path $commonPath)) {
    throw "Could not find $commonPath. Run this script from inside a full clone of the repository - it depends on the shared helper in Common/."
}

. $commonPath

#endregion Helpers ------------------------------------------------------------

Initialize-OutputPath -Path $OutputPath

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

        Connect-Site -Url $TenantAdminUrl -AppId $ClientId

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
Write-Host ''

$report = [System.Collections.Generic.List[PSCustomObject]]::new()

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

    $ownersFound = 0

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

        $ownersFound++
    }

    try {
        Connect-Site -Url $site -AppId $ClientId
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

        $siteTitle = $web.Title
        $template  = $web.WebTemplate
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

                $ownersFound++
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

                    $ownersFound++
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
            foreach ($owner in @(Get-PnPMicrosoft365GroupOwner -Identity $groupId -ErrorAction Stop)) {

                $login = [string]$owner.UserPrincipalName

                Add-OwnerRow -Site $site -SiteTitle $siteTitle -Template $template `
                             -GroupConnected $isGroupConnected -Source 'Microsoft365GroupOwner' `
                             -OwnerName $owner.DisplayName -OwnerLogin $login -OwnerEmail $owner.Email `
                             -PrincipalType 'User' `
                             -IsGuest ($login -imatch '#EXT#') `
                             -Note "Microsoft 365 group $groupId"

                $ownersFound++
            }
        }
        catch {
            # Reading group owners needs Graph permissions the SharePoint-only
            # connection may not carry; report rather than fail the whole run.
            Write-Warning "  Microsoft 365 group owners unavailable for $site : $($_.Exception.Message)"
            Add-OwnerRow -Site $site -SiteTitle $siteTitle -Template $template `
                         -GroupConnected $isGroupConnected -Source 'Microsoft365GroupOwner' `
                         -Note "Unavailable (Graph permission may be missing): $($_.Exception.Message)"
        }
    }

    if ($ownersFound -eq 0) {
        Add-OwnerRow -Site $site -SiteTitle $siteTitle -Template $template `
                     -GroupConnected $isGroupConnected -Source 'None' `
                     -Note 'No owners were found for this site'
    }

    Write-Host ("[{0}/{1}] {2} - {3} owner(s)" -f $counter, $sites.Count, $site, $ownersFound) -ForegroundColor DarkGray

    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }
}

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

$guestOwners = @($report | Where-Object { $_.IsGuest -eq $true -or $_.IsGuest -eq 'True' })

if ($guestOwners.Count -gt 0) {
    Write-Host "  Guest owners    : $($guestOwners.Count)" -ForegroundColor Yellow
}

Write-Host ''
Write-Host "  Report : $((Resolve-Path -Path $OutputPath).Path)" -ForegroundColor Green
