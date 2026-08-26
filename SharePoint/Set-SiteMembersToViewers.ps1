#Requires -Modules PnP.PowerShell

<#
.SYNOPSIS
    Moves members of a SharePoint site out of the Members group and into the
    Visitors group, demoting them from Edit to Read.

.DESCRIPTION
    By default only GUEST accounts are moved. Internal users are left alone unless
    -IncludeInternalUsers is specified.

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

.PARAMETER ClientId
    Client ID of the Entra app registration used by PnP.PowerShell. PnP no longer
    ships a shared multi-tenant app, so this is required unless you have set
    PnPManagementShellClientId in your environment. See the README.

.PARAMETER IncludeInternalUsers
    Also demote internal (non-guest) users. Off by default - guests only.

.PARAMETER IncludeSecurityGroups
    Also demote security groups and Microsoft 365 groups that appear as principals
    inside the Members group. Off by default; they are reported instead.

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

.EXAMPLE
    .\Set-SiteMembersToViewers.ps1 -SiteUrl https://contoso.sharepoint.com/sites/Project -ClientId $id -WhatIf

    Rehearsal against one site. Shows which guests would be demoted.

.EXAMPLE
    .\Set-SiteMembersToViewers.ps1 -SitesCsvPath .\sites.csv -ClientId $id

    Demote guests across every site listed in sites.csv.

.EXAMPLE
    .\Set-SiteMembersToViewers.ps1 -SiteUrl $url -ClientId $id -IncludeInternalUsers

    Demote everyone in Members, internal staff included.
#>

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Urls')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Urls')]
    [string[]]$SiteUrl,

    [Parameter(Mandatory, ParameterSetName = 'Csv')]
    [string]$SitesCsvPath,

    [string]$ClientId,

    [switch]$IncludeInternalUsers,

    [switch]$IncludeSecurityGroups,

    [string]$GuestLoginPattern = '(#ext#|urn:spo:guest)',

    [string[]]$ExcludeLogin = @(),

    [bool]$RemoveFromMembers = $true,

    [string]$OutputPath = ".\SharePoint_MembersToViewers_Log.csv",

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

function Test-IsGuest {
    param([Parameter(Mandatory)]$Principal, [Parameter(Mandatory)][string]$Pattern)

    $login = [string]$Principal.LoginName

    return ($login -imatch $Pattern)
}

# Shared CSV input handling - see Common/InputCsv.ps1
$commonPath = Join-Path -Path $PSScriptRoot -ChildPath '..\Common\InputCsv.ps1'

if (-not (Test-Path -Path $commonPath)) {
    throw "Could not find $commonPath. Run this script from inside a full clone of the repository - it depends on the shared helper in Common/."
}

. $commonPath

#endregion Helpers ------------------------------------------------------------

# Resolve the list of sites.
if ($PSCmdlet.ParameterSetName -eq 'Csv') {

    if (-not (Test-Path -Path $SitesCsvPath)) {
        throw "Sites CSV not found: $SitesCsvPath"
    }

    $csv = Import-InputCsv -Path $SitesCsvPath -Delimiter $Delimiter -RequiredColumns @('SiteUrl')

    if ($csv.Count -eq 0) { throw "Sites CSV is empty: $SitesCsvPath" }

    $sites = @($csv | Select-Object -ExpandProperty SiteUrl | Where-Object { $_ } | ForEach-Object { $_.Trim() })
}
else {
    $sites = @($SiteUrl | Where-Object { $_ } | ForEach-Object { $_.Trim() })
}

if ($sites.Count -eq 0) { throw 'No site URLs to process.' }

Initialize-OutputPath -Path $OutputPath

Write-Host ''
if ($IncludeInternalUsers) {
    Write-Host 'Scope: ALL members (guests and internal users).' -ForegroundColor Yellow
}
else {
    Write-Host 'Scope: guests only. Use -IncludeInternalUsers to include internal staff.' -ForegroundColor Cyan
}

if (-not $RemoveFromMembers) {
    Write-Host 'Copy mode: users will be added to Visitors but left in Members.' -ForegroundColor Yellow
}

Write-Host "Sites to process: $($sites.Count)" -ForegroundColor Cyan
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
        $connectParams = @{
            Url         = $site
            Interactive = $true
            ErrorAction = 'Stop'
        }

        if ($ClientId) { $connectParams['ClientId'] = $ClientId }

        Connect-PnPOnline @connectParams
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
        if ($type -and $type -ne 'User' -and -not $IncludeSecurityGroups) {
            Write-Result -Site $site -Title $title -Login $login -Email $email -PrincipalType $type `
                         -IsGuest $guest -Status 'Skipped' `
                         -Detail "Principal is a $type, not a user. Use -IncludeSecurityGroups to include it."
            $skipped++
            continue
        }

        if (-not $guest -and -not $IncludeInternalUsers) {
            Write-Result -Site $site -Title $title -Login $login -Email $email -PrincipalType $type `
                         -IsGuest $guest -Status 'Skipped' `
                         -Detail 'Internal user. Use -IncludeInternalUsers to include internal staff.'
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

    Write-Host "  Demoted: $moved   Left alone: $skipped" -ForegroundColor Green

    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }
}

Write-Progress -Activity 'Demoting site members' -Completed

$log | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -WhatIf:$false

Write-Host ''
Write-Host '--- Summary ---' -ForegroundColor Green

$log | Group-Object Status | Sort-Object Name | ForEach-Object {
    Write-Host ("  {0,-22} {1}" -f $_.Name, $_.Count)
}

Write-Host ''
Write-Host "  Log : $((Resolve-Path -Path $OutputPath).Path)" -ForegroundColor Green
