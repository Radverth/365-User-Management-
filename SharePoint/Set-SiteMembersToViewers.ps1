#Requires -Modules PnP.PowerShell

<#
.SYNOPSIS
    Moves members of a SharePoint site out of the Members group and into the
    Visitors group, demoting them from Edit to Read.

.DESCRIPTION
    Two things are demoted, because either alone leaves a hole: membership of the
    site's Members SharePoint group, and any permission granted to the person
    directly on the site. Moving someone out of Members achieves nothing if they
    also hold Edit directly. -SkipDirectPermissions limits it to group membership.

    -Scope decides who: guests only (the default), your own staff only, or both.

    For each site the script reads the associated Members group, works out who is
    in scope, adds them to the associated Visitors group and then removes them from
    Members. Adding happens before removing, so an interrupted run never leaves a
    user with no access at all.

    Safe to re-run: anyone already in Visitors is not re-added, and anyone already
    out of Members is not touched.

    Microsoft 365 group-connected sites (Teams sites): the people the SharePoint UI
    shows as "members" are members of the connected Microsoft 365 group, not of the
    SharePoint Members group. This script never touches Microsoft 365 group
    membership, because removing someone there also strips their Teams chat, group
    mailbox and calendar access. Those sites are still processed for anyone held
    directly in the SharePoint Members group, and the connected group is flagged in
    the output CSV so you can review it separately.

.PARAMETER SiteUrl
    One or more site collection URLs to process.

.PARAMETER SitesCsvPath
    CSV containing a SiteUrl column, as an alternative to -SiteUrl.

.PARAMETER AllSites
    Every site in the tenant. Requires -TenantAdminUrl.

    This demotes members across the whole tenant, so read the rehearsal output
    before committing. Personal OneDrive sites are excluded unless
    -IncludeOneDrive is given, and sites with no Visitors group are skipped rather
    than half-processed.

.PARAMETER TenantAdminUrl
    Your tenant admin URL, e.g. https://contoso-admin.sharepoint.com. Required
    with -AllSites.

.PARAMETER IncludeOneDrive
    With -AllSites, also include personal OneDrive sites. Excluded by default -
    demoting someone on their own OneDrive is rarely intended.

.PARAMETER ClientId
    Client ID of the Entra app registration used by PnP.PowerShell. PnP no longer
    ships a shared multi-tenant app, so this is required unless you have set
    PnPManagementShellClientId in your environment. See the README.

.PARAMETER Scope
    Who to demote:

      Guests  external people only (the default)
      Staff   your own users only, leaving guests alone
      Both    everyone

    Applies to SharePoint group membership and to direct permissions alike.

.PARAMETER SkipDirectPermissions
    Only fix group membership, leaving permissions granted directly on the site
    untouched. By default both are handled, because someone with direct Edit keeps
    editing no matter what group they are moved out of.

.PARAMETER IncludeInternalUsers
    Superseded by -Scope Both, and equivalent to it. Still accepted so existing
    commands keep working.

.PARAMETER IncludeSecurityGroups
    Also demote security groups and Microsoft 365 groups that appear as principals
    inside the Members group. Off by default; they are reported instead.

    On a Microsoft 365 group-connected site this is the important one. The Members
    group holds the connected group's member claim (shown in the UI as
    "<Site> Members"), so moving that single principal to Visitors demotes everyone
    in the team to read-only on the site at once - including people added to the
    group later. Microsoft 365 group membership itself is not touched, so Teams
    chat, the group mailbox and the calendar are unaffected.

.PARAMETER GuestLoginPattern
    Regex that identifies a guest by login name. The default matches both B2B
    guests (#ext#) and SharePoint-only external users (urn:spo:guest).

.PARAMETER ExcludeLogin
    Login names or email addresses to leave alone entirely. Useful for protecting
    service or break-glass accounts.

.PARAMETER RemoveFromMembers
    Whether to remove the user from Members after adding them to Visitors.
    Defaults to $true - a true move. Use -RemoveFromMembers:$false to copy into
    Visitors while leaving Members untouched, for a staged rollout.

.PARAMETER OutputPath
    CSV recording every principal considered and what happened to them.

.PARAMETER Delimiter
    Field separator of -SitesCsvPath. Detected automatically when omitted.

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
    .\Set-SiteMembersToViewers.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Project -ClientId $id -WhatIf

    Rehearsal against one site. Shows which guests would be demoted.

.EXAMPLE
    .\Set-SiteMembersToViewers.ps1 -SitesCsvPath .\sites.csv -ClientId $id

    Demote guests across every site listed in sites.csv.

.EXAMPLE
    .\Set-SiteMembersToViewers.ps1 -SiteUrl $url -ClientId $id -Scope Both

    Demote guests and staff alike, in groups and in direct permissions.
#>

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Urls')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Urls')]
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

    [ValidateSet('Guests', 'Staff', 'Both')]
    [string]$Scope,

    # Superseded by -Scope Both. Still honoured so existing commands keep working.
    [switch]$IncludeInternalUsers,

    [switch]$SkipDirectPermissions,

    [switch]$IncludeSecurityGroups,

    [string]$GuestLoginPattern = '(#ext#|urn:spo:guest)',

    [string[]]$ExcludeLogin = @(),

    [bool]$RemoveFromMembers = $true,

    [string]$OutputPath = ".\SharePoint_MembersToViewers_Log.csv",

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

# Permission levels that already mean read-only, plus the system-managed one that
# must never be stripped: Limited Access is what lets someone reach a parent so an
# item-level grant works, and removing it breaks that access.
$script:ReadOnlyRoles = @('Read', 'View Only', 'Restricted View', 'Restricted Read')
$script:SystemRoles   = @('Limited Access', 'Web-Only Limited Access')

function Get-ScopeDecision {
    <#  Decides whether one principal is in scope, and says why not when it is not.
        Shared by the group-membership pass and the direct-permission pass so the
        two cannot drift apart. #>
    param(
        [string]$PrincipalType,
        [bool]$IsGuest,
        [Parameter(Mandatory)][string]$Scope,
        [bool]$AllowGroups
    )

    # An unknown principal type is treated as a user, so the guests-only default
    # keeps protecting staff rather than falling through to the group branch.
    $isUser = (-not $PrincipalType) -or ($PrincipalType -eq 'User')

    if (-not $isUser) {

        if ($AllowGroups) { return [PSCustomObject]@{ InScope = $true; Reason = '' } }

        return [PSCustomObject]@{ InScope = $false
            Reason = "Principal is a $PrincipalType, not a user. Use -IncludeSecurityGroups to include it." }
    }

    switch ($Scope) {

        'Both' { return [PSCustomObject]@{ InScope = $true; Reason = '' } }

        'Guests' {
            if ($IsGuest) { return [PSCustomObject]@{ InScope = $true; Reason = '' } }

            return [PSCustomObject]@{ InScope = $false
                Reason = 'Your own user. Use -Scope Staff or -Scope Both to include them.' }
        }

        'Staff' {
            if (-not $IsGuest) { return [PSCustomObject]@{ InScope = $true; Reason = '' } }

            return [PSCustomObject]@{ InScope = $false
                Reason = 'Guest. Use -Scope Guests or -Scope Both to include guests.' }
        }
    }
}

function Test-IsGuest {
    param([Parameter(Mandatory)]$Principal, [Parameter(Mandatory)][string]$Pattern)

    $login = [string]$Principal.LoginName

    return ($login -imatch $Pattern)
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

# Built before the sites are resolved, because -AllSites has to sign in to the
# tenant admin site to enumerate them.
$auth = New-ScriptAuthContext -ClientId $ClientId -Tenant $Tenant -Thumbprint $Thumbprint `
                              -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword `
                              -NoPersistedLogin:$NoPersistedLogin

# Resolve the list of sites.
if ($PSCmdlet.ParameterSetName -eq 'All') {

    Write-Host "Connecting to $TenantAdminUrl to enumerate sites..." -ForegroundColor Cyan

    Connect-ScriptSite -Url $TenantAdminUrl -Auth $auth

    $tenantSites = @(Get-PnPTenantSite -ErrorAction Stop)

    if (-not $IncludeOneDrive) {
        $tenantSites = @($tenantSites | Where-Object { $_.Template -notlike 'SPSPERS*' })
    }

    $sites = @($tenantSites | Select-Object -ExpandProperty Url)

    Write-Host "  $($sites.Count) site(s) found." -ForegroundColor Green
    Write-Host ''
    Write-Warning "This will demote members across every one of those $($sites.Count) site(s)."
    Write-Warning 'Run it with -WhatIf first and read the log before committing.'
}
elseif ($PSCmdlet.ParameterSetName -eq 'Csv') {

    if (-not (Test-Path -Path $SitesCsvPath)) {
        throw "Sites CSV not found: $SitesCsvPath"
    }

    $csv = Import-InputCsv -Path $SitesCsvPath -Delimiter $Delimiter -Expected 'a site list with a SiteUrl column' -RequiredColumns @('SiteUrl')

    if ($csv.Count -eq 0) { throw "Sites CSV is empty: $SitesCsvPath" }

    $sites = @($csv | Select-Object -ExpandProperty SiteUrl | Where-Object { $_ } | ForEach-Object { $_.Trim() })
}
else {
    $sites = @($SiteUrl | Where-Object { $_ } | ForEach-Object { $_.Trim() })
}

if ($sites.Count -eq 0) { throw 'No site URLs to process.' }

Initialize-OutputPath -Path $OutputPath


Write-Host ''
# -Scope wins; the old switch maps onto it; neither means guests only.
if (-not $Scope) {
    $Scope = if ($IncludeInternalUsers) { 'Both' } else { 'Guests' }
}
elseif ($IncludeInternalUsers -and $Scope -ne 'Both') {
    Write-Warning "-IncludeInternalUsers is superseded by -Scope and was ignored; running with -Scope $Scope."
}

switch ($Scope) {
    'Both'   { Write-Host 'Scope: guests and staff.' -ForegroundColor Yellow }
    'Staff'  { Write-Host 'Scope: staff only. Guests keep their current access.' -ForegroundColor Yellow }
    'Guests' { Write-Host 'Scope: guests only. Use -Scope Staff or -Scope Both to include your own users.' -ForegroundColor Cyan }
}

if ($SkipDirectPermissions) {
    Write-Host 'Group membership only. Permissions granted directly on the site are left alone.' -ForegroundColor Yellow
}
else {
    Write-Host 'Covering both group membership and permissions granted directly on the site.' -ForegroundColor DarkGray
}

if (-not $RemoveFromMembers) {
    Write-Host 'Copy mode: users will be added to Visitors but left in Members.' -ForegroundColor Yellow
}

Write-Host "Sites to process: $($sites.Count)" -ForegroundColor Cyan
Write-PnPLoginAdvice -Auth $auth
Write-Host ''

$log = [System.Collections.Generic.List[PSCustomObject]]::new()

function Write-Result {
    param(
        [string]$Site,
        [string]$Title,
        [string]$Login,
        [string]$Email,
        [string]$PrincipalType = '',
        [string]$IsGuest = '',
        [Parameter(Mandatory)][string]$Status,
        [string]$Detail = ''
    )

    $log.Add([PSCustomObject]@{
        Timestamp     = (Get-Date).ToString('s')
        SiteUrl       = $Site
        PrincipalName = $Title
        LoginName     = $Login
        Email         = $Email
        PrincipalType = $PrincipalType
        IsGuest       = $IsGuest
        Status        = $Status
        Detail        = $Detail
    })
}

$siteCounter = 0

foreach ($site in $sites) {

    $siteCounter++

    Write-Progress -Activity 'Demoting site members' `
                   -Status "$siteCounter of $($sites.Count) - $site" `
                   -PercentComplete (($siteCounter / $sites.Count) * 100)

    Write-Host "[$siteCounter/$($sites.Count)] $site" -ForegroundColor Cyan

    # --- connect ------------------------------------------------------------

    try {
        Connect-ScriptSite -Url $site -Auth $auth
    }
    catch {
        Write-Warning "  Could not connect: $($_.Exception.Message)"
        Write-Result -Site $site -Status 'SiteFailed' -Detail "Connect failed: $($_.Exception.Message)"
        continue
    }

    # --- locate the associated groups ---------------------------------------

    try {
        $membersGroup  = Get-PnPGroup -AssociatedMemberGroup -ErrorAction Stop
        $visitorsGroup = Get-PnPGroup -AssociatedVisitorGroup -ErrorAction Stop
    }
    catch {
        Write-Warning "  Could not read the site's associated groups: $($_.Exception.Message)"
        Write-Result -Site $site -Status 'SiteFailed' -Detail "Associated groups unavailable: $($_.Exception.Message)"
        continue
    }

    if (-not $membersGroup) {
        Write-Warning '  This site has no associated Members group.'
        Write-Result -Site $site -Status 'SiteSkipped' -Detail 'No associated Members group'
        continue
    }

    if (-not $visitorsGroup) {
        # Without a Visitors group there is nowhere to demote people to. Removing
        # them from Members alone would silently revoke their access.
        Write-Warning '  This site has no associated Visitors group - nothing to demote into. Skipping.'
        Write-Result -Site $site -Status 'SiteSkipped' -Detail 'No associated Visitors group; no changes made'
        continue
    }

    # --- flag group-connected sites -----------------------------------------

    try {
        $pnpSite = Get-PnPSite -Includes GroupId -ErrorAction Stop
        $groupId = $pnpSite.GroupId

        if ($groupId -and $groupId -ne [Guid]::Empty) {
            Write-Host '  Microsoft 365 group-connected site. Group membership is NOT modified.' -ForegroundColor Yellow
            Write-Result -Site $site -Status 'Info' `
                         -Detail "Group-connected site (group $groupId). Members of the Microsoft 365 group keep Edit rights and were not touched."
        }
    }
    catch {
        Write-Verbose "  Could not determine group connection for $site : $($_.Exception.Message)"
    }

    # --- enumerate and filter -----------------------------------------------

    try {
        $members = @(Get-PnPGroupMember -Identity $membersGroup -ErrorAction Stop)
    }
    catch {
        Write-Warning "  Could not read members: $($_.Exception.Message)"
        Write-Result -Site $site -Status 'SiteFailed' -Detail "Member enumeration failed: $($_.Exception.Message)"
        continue
    }

    if ($members.Count -eq 0) {
        Write-Host '  Members group is empty.' -ForegroundColor DarkGray
        Write-Result -Site $site -Status 'SiteSkipped' -Detail 'Members group is empty'
        continue
    }

    # Read Visitors once so we can tell who is already there.
    try {
        $existingVisitors = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

        foreach ($visitor in @(Get-PnPGroupMember -Identity $visitorsGroup -ErrorAction Stop)) {
            [void]$existingVisitors.Add([string]$visitor.LoginName)
        }
    }
    catch {
        Write-Warning "  Could not read visitors: $($_.Exception.Message)"
        Write-Result -Site $site -Status 'SiteFailed' -Detail "Visitor enumeration failed: $($_.Exception.Message)"
        continue
    }

    $moved = 0
    $skipped = 0

    foreach ($principal in $members) {

        $login = [string]$principal.LoginName
        $title = [string]$principal.Title
        $email = [string]$principal.Email
        $type  = [string]$principal.PrincipalType
        $guest = Test-IsGuest -Principal $principal -Pattern $GuestLoginPattern

        # Explicit protection list wins over everything else.
        if ($ExcludeLogin -and ($ExcludeLogin -contains $login -or ($email -and $ExcludeLogin -contains $email))) {
            Write-Result -Site $site -Title $title -Login $login -Email $email -PrincipalType $type `
                         -IsGuest $guest -Status 'Excluded' -Detail 'Matched -ExcludeLogin'
            $skipped++
            continue
        }

        # Nested groups are principals too; demoting one moves everyone inside it.
        # On a Microsoft 365 group-connected site the Members group holds the
        # group's member claim, so this is the principal that governs the whole
        # team's access to the site.
        $decision = Get-ScopeDecision -PrincipalType $type -IsGuest $guest -Scope $Scope `
                                      -AllowGroups $IncludeSecurityGroups

        if (-not $decision.InScope) {
            Write-Result -Site $site -Title $title -Login $login -Email $email -PrincipalType $type `
                         -IsGuest $guest -Status 'Skipped' -Detail $decision.Reason
            $skipped++
            continue
        }

        # --- add to Visitors first ------------------------------------------

        $inVisitors = $existingVisitors.Contains($login)

        if ($inVisitors) {
            Write-Result -Site $site -Title $title -Login $login -Email $email -PrincipalType $type `
                         -IsGuest $guest -Status 'AlreadyVisitor' -Detail 'Already in the Visitors group'
        }
        elseif ($PSCmdlet.ShouldProcess("$title on $site", 'Add to Visitors')) {

            try {
                Add-PnPGroupMember -Identity $visitorsGroup -LoginName $login -ErrorAction Stop

                [void]$existingVisitors.Add($login)

                Write-Result -Site $site -Title $title -Login $login -Email $email -PrincipalType $type `
                             -IsGuest $guest -Status 'AddedToVisitors'
            }
            catch {
                # Do not remove from Members if the promotion to Visitors failed -
                # that would leave the user with no access to the site.
                Write-Result -Site $site -Title $title -Login $login -Email $email -PrincipalType $type `
                             -IsGuest $guest -Status 'Failed' `
                             -Detail "Add to Visitors failed, left in Members: $($_.Exception.Message)"

                Write-Warning "  $title : $($_.Exception.Message)"
                continue
            }
        }
        else {
            Write-Result -Site $site -Title $title -Login $login -Email $email -PrincipalType $type `
                         -IsGuest $guest -Status 'WhatIf' -Detail 'Would add to Visitors and remove from Members'
            continue
        }

        # --- then remove from Members ---------------------------------------

        if (-not $RemoveFromMembers) {
            Write-Result -Site $site -Title $title -Login $login -Email $email -PrincipalType $type `
                         -IsGuest $guest -Status 'KeptInMembers' -Detail '-RemoveFromMembers:$false was specified'
            $moved++
            continue
        }

        if ($PSCmdlet.ShouldProcess("$title on $site", 'Remove from Members')) {

            try {
                Remove-PnPGroupMember -Identity $membersGroup -LoginName $login -ErrorAction Stop

                Write-Result -Site $site -Title $title -Login $login -Email $email -PrincipalType $type `
                             -IsGuest $guest -Status 'RemovedFromMembers'

                $moved++
            }
            catch {
                Write-Result -Site $site -Title $title -Login $login -Email $email -PrincipalType $type `
                             -IsGuest $guest -Status 'Failed' `
                             -Detail "Now in Visitors but removal from Members failed: $($_.Exception.Message)"

                Write-Warning "  $title : $($_.Exception.Message)"
            }
        }
    }

    # --- permissions granted directly on the site -------------------------

    # Moving someone between SharePoint groups achieves nothing if they also hold
    # Edit directly on the site, so unless told otherwise the direct grants are
    # reduced to Read as well.
    if (-not $SkipDirectPermissions) {

        $assignments = @()

        try {
            $web         = Get-PnPWeb -Includes RoleAssignments -ErrorAction Stop
            $assignments = @($web.RoleAssignments)
        }
        catch {
            Write-Warning "  Could not read direct permissions: $($_.Exception.Message)"
            Write-Result -Site $site -Status 'Failed' `
                         -Detail "Direct permissions could not be read: $($_.Exception.Message)"
        }

        foreach ($assignment in $assignments) {

            try {
                $member   = Get-PnPProperty -ClientObject $assignment -Property Member -ErrorAction Stop
                $bindings = @(Get-PnPProperty -ClientObject $assignment -Property RoleDefinitionBindings -ErrorAction Stop)
            }
            catch {
                Write-Result -Site $site -Status 'Failed' `
                             -Detail "A direct permission entry could not be read: $($_.Exception.Message)"
                continue
            }

            $dLogin = [string]$member.LoginName
            $dTitle = [string]$member.Title
            $dType  = [string]$member.PrincipalType

            # The site's own SharePoint groups are the membership pass's business.
            if ($dType -eq 'SharePointGroup') { continue }

            # Never a person, and removing its access breaks the site.
            if ($dLogin -eq 'SHAREPOINT\system') { continue }

            $dEmail = [string]$member.Email
            $dGuest = ($dLogin -imatch $GuestLoginPattern)

            if ($ExcludeLogin -and ($ExcludeLogin -contains $dLogin -or ($dEmail -and $ExcludeLogin -contains $dEmail))) {
                Write-Result -Site $site -Title $dTitle -Login $dLogin -Email $dEmail -PrincipalType $dType `
                             -IsGuest $dGuest -Status 'Excluded' -Detail 'Direct permission; matched -ExcludeLogin'
                $skipped++
                continue
            }

            $dDecision = Get-ScopeDecision -PrincipalType $dType -IsGuest $dGuest -Scope $Scope `
                                           -AllowGroups $IncludeSecurityGroups

            if (-not $dDecision.InScope) {
                Write-Result -Site $site -Title $dTitle -Login $dLogin -Email $dEmail -PrincipalType $dType `
                             -IsGuest $dGuest -Status 'Skipped' -Detail "Direct permission; $($dDecision.Reason)"
                $skipped++
                continue
            }

            $roles = @($bindings | ForEach-Object { [string]$_.Name })

            # Anything that is not already read-only, and is not system-managed.
            $editRoles = @($roles | Where-Object {
                $_ -notin $script:ReadOnlyRoles -and $_ -notin $script:SystemRoles
            })

            if ($editRoles.Count -eq 0) {
                Write-Result -Site $site -Title $dTitle -Login $dLogin -Email $dEmail -PrincipalType $dType `
                             -IsGuest $dGuest -Status 'AlreadyReadOnly' `
                             -Detail "Direct permission is already read-only: $($roles -join ', ')"
                continue
            }

            if ($PSCmdlet.ShouldProcess("$dTitle on $site", "Reduce direct permission to Read")) {

                try {
                    $permissionParams = @{
                        User        = $dLogin
                        RemoveRole  = $editRoles
                        ErrorAction = 'Stop'
                    }

                    # Only grant Read if they do not already hold it alongside the
                    # edit-level role being removed.
                    if ($roles -notcontains 'Read') { $permissionParams['AddRole'] = 'Read' }

                    Set-PnPWebPermission @permissionParams

                    Write-Result -Site $site -Title $dTitle -Login $dLogin -Email $dEmail -PrincipalType $dType `
                                 -IsGuest $dGuest -Status 'DirectPermissionReduced' `
                                 -Detail "Removed $($editRoles -join ', '); left with Read"

                    $moved++
                }
                catch {
                    Write-Result -Site $site -Title $dTitle -Login $dLogin -Email $dEmail -PrincipalType $dType `
                                 -IsGuest $dGuest -Status 'Failed' `
                                 -Detail "Direct permission unchanged ($($editRoles -join ', ')): $($_.Exception.Message)"

                    Write-Warning "  $dTitle : direct permission unchanged - $($_.Exception.Message)"
                }
            }
            else {
                Write-Result -Site $site -Title $dTitle -Login $dLogin -Email $dEmail -PrincipalType $dType `
                             -IsGuest $dGuest -Status 'WhatIf' `
                             -Detail "Would reduce direct permission from $($editRoles -join ', ') to Read"
            }
        }
    }

    Write-Host "  Demoted: $moved   Left alone: $skipped" -ForegroundColor Green

}

# One disconnect at the end. Doing it per site drops the token context and makes
# the next Connect prompt again. The persisted cache is left intact - clear it
# with Disconnect-PnPOnline -ClearPersistedLogin.
try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }

Write-Progress -Activity 'Demoting site members' -Completed

$log | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -WhatIf:$false

Write-Host ''
Write-Host '--- Summary ---' -ForegroundColor Green

$log | Group-Object Status | Sort-Object Name | ForEach-Object {
    Write-Host ("  {0,-22} {1}" -f $_.Name, $_.Count)
}

Write-Host ''
Write-Host "  Log : $((Resolve-Path -Path $OutputPath).Path)" -ForegroundColor Green
