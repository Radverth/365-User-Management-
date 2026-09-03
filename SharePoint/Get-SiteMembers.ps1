#Requires -Modules PnP.PowerShell

<#
.SYNOPSIS
    Reports who has member-level access to SharePoint sites, to CSV.

.DESCRIPTION
    Read-only. Nothing is modified. The companion to Get-SiteOwners.ps1: that one
    answers "who runs this site", this one answers "who can get into it".

    Access reaches a site by several routes, and a report that only reads the
    Members group misses most of them. All of these are collected, each row tagged
    with where it came from:

      MembersGroup            The site's associated Members group.
      VisitorsGroup           The associated Visitors group - read-only access,
                              but access.
      SharePointGroup         Any other SharePoint group on the site. Real sites
                              accumulate these and they are easily forgotten.
      DirectPermission        People and security groups given permission on the
                              site itself rather than through a group.
      Microsoft365GroupMember Members of the connected Microsoft 365 group, for
                              group-connected (Teams) sites. These people have
                              access even when the Members group looks empty.

    Owners are left to Get-SiteOwners.ps1. Add -IncludeOwnersGroup to have them
    here too, when you want one file covering everybody.

    Each person produces one row per route, so somebody who is in the Members
    group and also holds a direct permission appears twice. That is the point: it
    is what you have to unpick to remove their access.

    The permission levels behind each row are reported in the Roles column, read
    from the site's own role assignments, so a custom group with Full Control does
    not read as ordinary membership.

.PARAMETER SiteUrl
    One or more site collection URLs to report on.

.PARAMETER SitesCsvPath
    CSV containing a SiteUrl column, as an alternative to -SiteUrl.

.PARAMETER AllSites
    Report on every site in the tenant. Requires -TenantAdminUrl and SharePoint
    administrator rights.

    Each site is opened individually to read its groups, so allow time on a large
    tenant. Being SharePoint Administrator does not make you a site collection
    administrator everywhere; sites that refuse to open produce an Error row
    saying so rather than being silently absent.

.PARAMETER TenantAdminUrl
    Your tenant admin URL, e.g. https://contoso-admin.sharepoint.com. Required with
    -AllSites.

.PARAMETER IncludeOneDrive
    With -AllSites, also include personal OneDrive sites. Excluded by default.

.PARAMETER ClientId
    Client ID of the Entra app registration used by PnP.PowerShell.

.PARAMETER IncludeMembersGroup
    Include the associated Members group. Defaults to $true.

.PARAMETER IncludeVisitorsGroup
    Include the associated Visitors group. Defaults to $true.

.PARAMETER IncludeOtherGroups
    Include every other SharePoint group on the site. Defaults to $true.

.PARAMETER IncludeDirectPermissions
    Include principals given permission on the site directly. Defaults to $true.

.PARAMETER IncludeMicrosoft365GroupMembers
    Include members of the connected Microsoft 365 group. Defaults to $true.

.PARAMETER IncludeOwnersGroup
    Also include the associated Owners group, for a single file covering everyone.
    Off by default, because Get-SiteOwners.ps1 reports owners properly - including
    site collection administrators, which this script does not look at.

.PARAMETER GuestsOnly
    Report only guests. Useful before a tenant migration, when the question is
    which external people can reach which sites.

.PARAMETER GuestLoginPattern
    Regular expression that identifies a guest login. Defaults to the two forms
    SharePoint uses.

.PARAMETER IncludeSystemPrincipals
    Keep principals that are not real people: the SHAREPOINT\system account, the
    Everyone and Everyone except external users claims, and the tenant-wide
    administrator role claims. Excluded by default, and the number removed is
    reported at the end.

    Worth turning on at least once: "Everyone except external users" on a site is
    a finding, not noise.

.PARAMETER OutputPath
    Destination CSV.

.PARAMETER Delimiter
    Field separator of -SitesCsvPath. Detected automatically when omitted.

.PARAMETER Tenant
    Tenant name, e.g. contoso.onmicrosoft.com. Required for app-only sign-in.

.PARAMETER Thumbprint
    Thumbprint of a certificate in your certificate store. Supplying it switches
    the script to app-only sign-in, where access comes from the app registration's
    application permissions rather than your own site access.

.PARAMETER CertificatePath
    Path to a .pfx instead of -Thumbprint. Also switches to app-only sign-in.

.PARAMETER CertificatePassword
    Password for -CertificatePath, as a SecureString.

.PARAMETER NoPersistedLogin
    Sign in afresh for every site instead of reusing a cached token.

.EXAMPLE
    .\Get-SiteMembers.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Project -ClientId $id

.EXAMPLE
    .\Get-SiteMembers.ps1 -AllSites -TenantAdminUrl https://contoso-admin.sharepoint.com -ClientId $id

.EXAMPLE
    .\Get-SiteMembers.ps1 -AllSites -TenantAdminUrl https://contoso-admin.sharepoint.com -ClientId $id -GuestsOnly

    Every guest with access to any site, and how they got it.

.EXAMPLE
    .\Get-SiteMembers.ps1 -SitesCsvPath .\sites.csv -ClientId $id -OutputPath .\members.csv
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

    [bool]$IncludeMembersGroup = $true,

    [bool]$IncludeVisitorsGroup = $true,

    [bool]$IncludeOtherGroups = $true,

    [bool]$IncludeDirectPermissions = $true,

    [bool]$IncludeMicrosoft365GroupMembers = $true,

    [switch]$IncludeOwnersGroup,

    [switch]$GuestsOnly,

    [string]$GuestLoginPattern = '(#ext#|urn:spo:guest)',

    [switch]$IncludeSystemPrincipals,

    [string]$OutputPath = ".\SharePoint_SiteMembers.csv",

    [string]$Delimiter,

    [string]$Tenant,

    [string]$Thumbprint,

    [string]$CertificatePath,

    [System.Security.SecureString]$CertificatePassword,

    [switch]$NoPersistedLogin
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
    throw "Could not find $commonPath. Run this script from inside the complete toolkit folder - it depends on the shared helper in Common/."
}

. $commonPath

$pnpConnectPath = Join-Path -Path $PSScriptRoot -ChildPath '..\Common\PnPConnect.ps1'

if (-not (Test-Path -Path $pnpConnectPath)) {
    throw "Could not find $pnpConnectPath. Run this script from inside the complete toolkit folder - it depends on the shared helpers in Common/."
}

. $pnpConnectPath

#endregion Helpers ------------------------------------------------------------

Initialize-OutputPath -Path $OutputPath

$auth = New-ScriptAuthContext -ClientId $ClientId -Tenant $Tenant -Thumbprint $Thumbprint `
                              -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword `
                              -NoPersistedLogin:$NoPersistedLogin

# Resolve the list of sites.
$sites = @()

# Url (lowercased, no trailing slash) -> the Get-PnPTenantSite record, when
# -AllSites was used. Empty in the other modes.
$tenantSiteIndex = @{}

switch ($PSCmdlet.ParameterSetName) {

    'Csv' {
        if (-not (Test-Path -Path $SitesCsvPath)) { throw "Sites CSV not found: $SitesCsvPath" }

        $csv = Import-InputCsv -Path $SitesCsvPath -Delimiter $Delimiter -Expected 'a site list with a SiteUrl column' -RequiredColumns @('SiteUrl')

        if ($csv.Count -eq 0) { throw "Sites CSV is empty: $SitesCsvPath" }

        $sites = @($csv | Select-Object -ExpandProperty SiteUrl | Where-Object { $_ } | ForEach-Object { $_.Trim() })
    }

    'All' {
        # Correct the address before connecting. Pointed at an ordinary site,
        # Get-PnPTenantSite fails with a message about content types that says
        # nothing about the real mistake.
        $admin = Resolve-TenantAdminUrl -Url $TenantAdminUrl

        if ($admin.Corrected) {
            Write-Host "  $($admin.Original) is not the admin address - using $($admin.Url)" -ForegroundColor Yellow
        }

        Write-Host "Connecting to $($admin.Url) to enumerate sites..." -ForegroundColor Cyan

        Connect-ScriptSite -Url $admin.Url -Auth $auth

        $tenantSites = @(Get-ScriptTenantSite -AdminUrl $admin.Url -Auth $auth)

        if (-not $IncludeOneDrive) {
            $tenantSites = @($tenantSites | Where-Object { $_.Template -notlike 'SPSPERS*' })
        }

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
    [void]$usage.AppendLine('    .\Get-SiteMembers.ps1 -AllSites -TenantAdminUrl https://<tenant>-admin.sharepoint.com -ClientId <id>')
    [void]$usage.AppendLine('')
    [void]$usage.AppendLine('  Specific sites:')
    [void]$usage.AppendLine('    .\Get-SiteMembers.ps1 -SiteUrl https://<tenant>.sharepoint.com/sites/One,https://<tenant>.sharepoint.com/sites/Two -ClientId <id>')
    [void]$usage.AppendLine('')
    [void]$usage.AppendLine('  From a CSV with a SiteUrl column:')
    [void]$usage.AppendLine('    .\Get-SiteMembers.ps1 -SitesCsvPath .\sites.csv -ClientId <id>')

    throw $usage.ToString()
}

Write-Host "Reporting members for $($sites.Count) site(s)..." -ForegroundColor Cyan
Write-PnPLoginAdvice -Auth $auth
Write-Host ''

$report = [System.Collections.Generic.List[PSCustomObject]]::new()

$script:SeenRows                 = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$script:SiteRowCount             = 0
$script:FilteredSystemPrincipals = 0
$script:FilteredNonGuests        = 0

# Claims that are not people. Everyone and Everyone except external users are the
# two that matter most, because a site carrying either is open far wider than its
# group membership suggests.
$script:SystemLogins = @(
    'SHAREPOINT\system'
    'c:0(.s|true'                       # Everyone
    'c:0-.f|rolemanager|spo-grid-all-users'
)

function Test-IsSystemPrincipal {
    param([string]$Login)

    if (-not $Login) { return $false }

    if ($Login -in $script:SystemLogins) { return $true }

    # Tenant-wide administrator role claims, identical on every site.
    if ($Login -like 'c:0t.c|tenant|*') { return $true }

    # Everyone except external users, whose claim carries the tenant id.
    if ($Login -like 'c:0-.f|rolemanager|spo-grid-all-users*') { return $true }

    return $false
}

function Add-MemberRow {
    param(
        [Parameter(Mandatory)][string]$Site,
        [string]$SiteTitle,
        [string]$Template,
        [string]$GroupConnected,
        [Parameter(Mandatory)][string]$Source,
        [string]$GroupName,
        [string]$MemberName,
        [string]$MemberLogin,
        [string]$MemberEmail,
        [string]$PrincipalType,
        [string]$Roles,
        [string]$Note = ''
    )

    $isGuest = [bool]($MemberLogin -and $MemberLogin -imatch $GuestLoginPattern)

    # An Error or None row carries no principal, so it is never filtered away -
    # losing the reason a site produced nothing would be the worst outcome here.
    $isPrincipalRow = [bool]($MemberName -or $MemberLogin)

    if ($isPrincipalRow) {

        if (-not $IncludeSystemPrincipals -and (Test-IsSystemPrincipal -Login $MemberLogin)) {
            $script:FilteredSystemPrincipals++
            return
        }

        if ($GuestsOnly -and -not $isGuest) {
            $script:FilteredNonGuests++
            return
        }
    }

    # The same person can surface twice from one route - a repeated Graph page, or
    # a group listed under two names.
    $key = "$Site|$Source|$GroupName|$MemberLogin|$MemberName|$Note"

    if (-not $script:SeenRows.Add($key)) { return }

    if ($isPrincipalRow) { $script:SiteRowCount++ }

    $report.Add([PSCustomObject]@{
        SiteUrl          = $Site
        SiteTitle        = $SiteTitle
        Template         = $Template
        IsGroupConnected = $GroupConnected
        MemberSource     = $Source
        GroupName        = $GroupName
        MemberName       = $MemberName
        MemberLogin      = $MemberLogin
        MemberEmail      = $MemberEmail
        PrincipalType    = $PrincipalType
        Roles            = $Roles
        IsGuest          = $isGuest
        Note             = $Note
    })
}

$counter = 0

foreach ($site in $sites) {

    $counter++

    Write-Progress -Activity 'Collecting site members' `
                   -Status "$counter of $($sites.Count) - $site" `
                   -PercentComplete (($counter / $sites.Count) * 100)

    $tenantRecord = $tenantSiteIndex[$site.TrimEnd('/').ToLowerInvariant()]

    $siteTitle = if ($tenantRecord) { [string]$tenantRecord.Title } else { '' }
    $template  = if ($tenantRecord) { [string]$tenantRecord.Template } else { '' }
    $groupId   = $null

    if ($tenantRecord -and $tenantRecord.GroupId -and $tenantRecord.GroupId -ne [Guid]::Empty) {
        $groupId = $tenantRecord.GroupId
    }

    $script:SiteRowCount = 0

    try {
        Connect-ScriptSite -Url $site -Auth $auth
    }
    catch {
        Write-Warning "Could not open $site : $($_.Exception.Message)"

        Add-MemberRow -Site $site -SiteTitle $siteTitle -Template $template `
                      -GroupConnected ([bool]$groupId) -Source 'Error' `
                      -Note "Could not open the site, so nothing could be read: $($_.Exception.Message)"
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

    # --- role assignments, read once ----------------------------------------

    # One pass over the site's role assignments does two jobs: it produces the
    # direct-permission rows, and it records what each SharePoint group is
    # actually allowed to do, so a custom group with Full Control is not reported
    # as ordinary membership.
    $groupRoles       = @{}
    $directAssignments = [System.Collections.Generic.List[object]]::new()

    try {
        $webWithRoles = Get-PnPWeb -Includes RoleAssignments -ErrorAction Stop

        foreach ($assignment in @($webWithRoles.RoleAssignments)) {

            try {
                $principal = Get-PnPProperty -ClientObject $assignment -Property Member -ErrorAction Stop
                $bindings  = @(Get-PnPProperty -ClientObject $assignment -Property RoleDefinitionBindings -ErrorAction Stop)
            }
            catch {
                Add-MemberRow -Site $site -SiteTitle $siteTitle -Template $template `
                              -GroupConnected $isGroupConnected -Source 'Error' `
                              -Note "A permission entry could not be read: $($_.Exception.Message)"
                continue
            }

            # Limited Access is plumbing - SharePoint grants it so a person can
            # reach one item in a library - and reporting it as site access is
            # misleading.
            $roleNames = @($bindings |
                            ForEach-Object { [string]$_.Name } |
                            Where-Object { $_ -and $_ -notin @('Limited Access', 'Web-Only Limited Access') })

            if ($roleNames.Count -eq 0) { continue }

            if ([string]$principal.PrincipalType -eq 'SharePointGroup') {
                $groupRoles[[string]$principal.Title] = ($roleNames -join '; ')
            }
            else {
                $directAssignments.Add([pscustomobject]@{ Principal = $principal; Roles = ($roleNames -join '; ') })
            }
        }
    }
    catch {
        Write-Warning "  Could not read permissions for $site : $($_.Exception.Message)"

        Add-MemberRow -Site $site -SiteTitle $siteTitle -Template $template `
                      -GroupConnected $isGroupConnected -Source 'Error' `
                      -Note "Site permissions could not be read, so roles and direct permissions are missing: $($_.Exception.Message)"
    }

    # --- SharePoint groups ---------------------------------------------------

    # Title -> the label this report gives it. Built from the associated groups so
    # the Members group is called MembersGroup rather than by whatever it was
    # renamed to, and so it is not reported a second time as a custom group.
    $associated = [ordered]@{}

    foreach ($pair in @(
        @{ Switch = $IncludeMembersGroup;             Param = 'AssociatedMemberGroup';  Source = 'MembersGroup' }
        @{ Switch = $IncludeVisitorsGroup;            Param = 'AssociatedVisitorGroup'; Source = 'VisitorsGroup' }
        @{ Switch = [bool]$IncludeOwnersGroup;        Param = 'AssociatedOwnerGroup';   Source = 'OwnersGroup' }
    )) {

        # -AssociatedMemberGroup and friends are switches, so the one wanted is
        # built into a splat rather than branched three ways.
        $groupSplat = @{ $pair.Param = $true; ErrorAction = 'Stop' }

        try   { $group = Get-PnPGroup @groupSplat }
        catch { $group = $null }

        if ($group) { $associated[[string]$group.Title] = @{ Group = $group; Source = $pair.Source; Wanted = $pair.Switch } }
    }

    $groupsToRead = [System.Collections.Generic.List[object]]::new()

    foreach ($title in $associated.Keys) {
        if ($associated[$title].Wanted) {
            $groupsToRead.Add([pscustomobject]@{ Group = $associated[$title].Group; Source = $associated[$title].Source })
        }
    }

    if ($IncludeOtherGroups) {
        try {
            foreach ($group in @(Get-PnPGroup -ErrorAction Stop)) {

                if ($associated.Contains([string]$group.Title)) { continue }

                $groupsToRead.Add([pscustomobject]@{ Group = $group; Source = 'SharePointGroup' })
            }
        }
        catch {
            Write-Warning "  Could not list SharePoint groups for $site : $($_.Exception.Message)"

            Add-MemberRow -Site $site -SiteTitle $siteTitle -Template $template `
                          -GroupConnected $isGroupConnected -Source 'Error' `
                          -Note "SharePoint groups could not be listed: $($_.Exception.Message)"
        }
    }

    foreach ($entry in $groupsToRead) {

        $group     = $entry.Group
        $groupName = [string]$group.Title
        $roles     = [string]$groupRoles[$groupName]

        try {
            $members = @(Get-PnPGroupMember -Identity $group -ErrorAction Stop)
        }
        catch {
            Write-Warning "  Could not read $groupName on $site : $($_.Exception.Message)"

            Add-MemberRow -Site $site -SiteTitle $siteTitle -Template $template `
                          -GroupConnected $isGroupConnected -Source $entry.Source -GroupName $groupName `
                          -Roles $roles -Note "Group members could not be read: $($_.Exception.Message)"
            continue
        }

        foreach ($member in $members) {

            Add-MemberRow -Site $site -SiteTitle $siteTitle -Template $template `
                          -GroupConnected $isGroupConnected -Source $entry.Source -GroupName $groupName `
                          -MemberName $member.Title -MemberLogin $member.LoginName -MemberEmail $member.Email `
                          -PrincipalType $member.PrincipalType -Roles $roles
        }
    }

    # --- direct permissions --------------------------------------------------

    if ($IncludeDirectPermissions) {

        foreach ($entry in $directAssignments) {

            $principal = $entry.Principal

            Add-MemberRow -Site $site -SiteTitle $siteTitle -Template $template `
                          -GroupConnected $isGroupConnected -Source 'DirectPermission' `
                          -MemberName $principal.Title -MemberLogin $principal.LoginName -MemberEmail $principal.Email `
                          -PrincipalType $principal.PrincipalType -Roles $entry.Roles `
                          -Note 'Permission on the site itself, not through a group'
        }
    }

    # --- Microsoft 365 group members ----------------------------------------

    if ($IncludeMicrosoft365GroupMembers -and $groupId) {
        try {
            $groupMembers = @(Get-PnPMicrosoft365GroupMember -Identity $groupId -ErrorAction Stop)

            foreach ($member in $groupMembers) {

                $memberName  = [string]$member.DisplayName
                $memberLogin = [string]$member.UserPrincipalName
                $memberMail  = [string]$member.Mail

                if (-not $memberMail)  { $memberMail  = [string]$member.Email }
                if (-not $memberLogin) { $memberLogin = $memberMail }

                $note = ''

                # Same shortcoming as the owner cmdlet: sometimes only the
                # directory object ID comes back. A row identifying somebody by ID
                # beats a blank line, as long as it says why.
                if (-not $memberName -and -not $memberLogin) {

                    $memberLogin = [string]$member.Id

                    if ($memberLogin) {
                        $note = 'Only the directory object ID was returned - grant the app User.Read.All to resolve names'
                    }
                }

                Add-MemberRow -Site $site -SiteTitle $siteTitle -Template $template `
                              -GroupConnected $isGroupConnected -Source 'Microsoft365GroupMember' `
                              -GroupName "Microsoft 365 group $groupId" `
                              -MemberName $memberName -MemberLogin $memberLogin -MemberEmail $memberMail `
                              -PrincipalType 'User' -Roles 'Member of the connected group' -Note $note
            }
        }
        catch {
            $reason = $_.Exception.Message

            if ($reason -match 'does not exist|Not Found|\(404\)') {
                Write-Warning "  $site is connected to Microsoft 365 group $groupId, which no longer exists."

                Add-MemberRow -Site $site -SiteTitle $siteTitle -Template $template `
                              -GroupConnected $isGroupConnected -Source 'OrphanedGroup' `
                              -Note "Connected Microsoft 365 group $groupId no longer exists, so its members could not be read."
            }
            else {
                Write-Warning "  Microsoft 365 group members unavailable for $site : $reason"

                Add-MemberRow -Site $site -SiteTitle $siteTitle -Template $template `
                              -GroupConnected $isGroupConnected -Source 'Microsoft365GroupMember' `
                              -Note "Unavailable (Graph permission may be missing): $reason"
            }
        }
    }

    if ($script:SiteRowCount -eq 0) {
        Add-MemberRow -Site $site -SiteTitle $siteTitle -Template $template `
                      -GroupConnected $isGroupConnected -Source 'None' `
                      -Note $(if ($GuestsOnly) { 'No guests have access to this site' } else { 'Nobody was found with member-level access to this site' })
    }

    Write-Host ("[{0}/{1}] {2} - {3} row(s)" -f $counter, $sites.Count, $site, $script:SiteRowCount) -ForegroundColor DarkGray
}

# One disconnect at the end. Doing it per site drops the token context and makes
# the next Connect prompt again.
try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }

Write-Progress -Activity 'Collecting site members' -Completed

$report |
    Sort-Object SiteUrl, MemberSource, GroupName, MemberName |
    Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host ''
Write-Host '--- Summary ---' -ForegroundColor Green
Write-Host "  Sites processed : $($sites.Count)"
Write-Host "  Rows            : $($report.Count)"

$report | Group-Object MemberSource | Sort-Object Name | ForEach-Object {
    Write-Host ("    {0,-26} {1}" -f $_.Name, $_.Count)
}

$people = @($report | Where-Object { $_.MemberLogin -or $_.MemberName })

$distinct = @($people | Where-Object { $_.MemberLogin } |
                Select-Object -ExpandProperty MemberLogin -Unique)

Write-Host "  Distinct people : $($distinct.Count)"

$guests = @($people | Where-Object { $_.IsGuest -eq $true -or $_.IsGuest -eq 'True' })

if ($guests.Count -gt 0) {
    Write-Host "  Guest rows      : $($guests.Count)" -ForegroundColor Yellow
}

if ($script:FilteredNonGuests -gt 0) {
    Write-Host "  Filtered out    : $($script:FilteredNonGuests) internal principal(s), because -GuestsOnly was used." -ForegroundColor DarkGray
}

if ($script:FilteredSystemPrincipals -gt 0) {
    Write-Host "  Filtered out    : $($script:FilteredSystemPrincipals) system / tenant-admin-role principal(s). Use -IncludeSystemPrincipals to keep them." -ForegroundColor DarkGray
}

$orphaned = @($report | Where-Object { $_.MemberSource -eq 'OrphanedGroup' })

if ($orphaned.Count -gt 0) {
    Write-Host "  Orphaned sites  : $($orphaned.Count) connected to a Microsoft 365 group that no longer exists." -ForegroundColor Yellow
}

Write-Host ''
Write-Host "  Report : $((Resolve-Path -Path $OutputPath).Path)" -ForegroundColor Green
