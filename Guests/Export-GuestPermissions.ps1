#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Identity.DirectoryManagement

<#
.SYNOPSIS
    Exports every guest user in the source tenant along with their group memberships.

.DESCRIPTION
    Run this against tenant A. It produces a CSV that Import-GuestPermissions.ps1
    replays against tenant B.

    Only DIRECT group memberships are exported. Nested (transitive) memberships are
    deliberately excluded because they are inherited automatically once the direct
    membership is recreated in the target tenant - exporting them would create
    duplicate, incorrect direct memberships.

    Groups that cannot be populated through Microsoft Graph are still exported, but
    flagged Importable = False and written to a second CSV for manual handling:
      - Distribution lists and mail-enabled security groups (Exchange Online only)
      - Dynamic groups (membership is governed by a rule, not by member objects)

.PARAMETER OutputPath
    CSV of guests and their memberships. Consumed by Import-GuestPermissions.ps1.

.PARAMETER UnsupportedOutputPath
    CSV containing only the memberships that Graph cannot recreate.

.PARAMETER TenantId
    Optional tenant ID or domain to sign in against, useful when your account exists
    in more than one tenant.

.PARAMETER IncludeDisabledGuests
    Include guests whose account is blocked from sign-in. Excluded by default.

.PARAMETER IncludePendingAcceptance
    Include guests who have never redeemed their original invitation
    (externalUserState = PendingAcceptance). Excluded by default.

.PARAMETER ExcludeGroupOwnership
    Skip the lookup of groups each guest owns. Speeds up the export on large tenants.

.EXAMPLE
    .\Export-GuestPermissions.ps1

.EXAMPLE
    .\Export-GuestPermissions.ps1 -OutputPath C:\Migration\GuestsA.csv -IncludeDisabledGuests
#>

[CmdletBinding()]
param(
    [string]$OutputPath = ".\TenantA_GuestPermissions.csv",

    [string]$UnsupportedOutputPath = ".\TenantA_GuestPermissions_Unsupported.csv",

    [string]$TenantId,

    [switch]$IncludeDisabledGuests,

    [switch]$IncludePendingAcceptance,

    [switch]$ExcludeGroupOwnership
)

$ErrorActionPreference = 'Stop'

#region Helpers ---------------------------------------------------------------

function Initialize-OutputPath {
    <#  Creates the parent folder when the path has one. A bare file name has no
        parent, so Split-Path returns an empty string and New-Item must be skipped. #>
    param([Parameter(Mandatory)][string]$Path)

    $parent = Split-Path -Path $Path -Parent

    if ($parent -and -not (Test-Path -Path $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
}

function Resolve-ExternalEmail {
    <#  Returns the guest's real (home tenant) email address.

        This is the only identifier that survives a tenant migration. A guest's UPN
        is tenant-specific - john@contoso.com becomes
        john_contoso.com#EXT#@tenanta.onmicrosoft.com in tenant A and
        john_contoso.com#EXT#@tenantb.onmicrosoft.com in tenant B - so the UPN
        cannot be used to match the same person across tenants.

        Mail is preferred, then otherMails, then the UPN is decoded as a fallback. #>
    param([Parameter(Mandatory)]$User)

    if ($User.Mail) { return $User.Mail.Trim() }

    if ($User.OtherMails -and $User.OtherMails.Count -gt 0) {
        return ($User.OtherMails | Select-Object -First 1).Trim()
    }

    $upn = $User.UserPrincipalName

    if ($upn -and $upn -match '^(?<local>.+)#EXT#@') {
        # john_contoso.com#EXT#@tenant.onmicrosoft.com -> john@contoso.com
        # Only the LAST underscore is the encoded '@'; the local part may contain others.
        $local = $Matches['local']
        $split = $local.LastIndexOf('_')

        if ($split -gt 0) {
            return ($local.Substring(0, $split) + '@' + $local.Substring($split + 1))
        }
    }

    return $null
}

function Get-GroupFacts {
    <#  Normalises a directory object returned by Get-MgUserMemberOf into the fields
        the import needs, and decides whether Graph can recreate the membership.

        Results are cached by object ID: a group shared by 200 guests is only
        classified once instead of 200 times. #>
    param(
        [Parameter(Mandatory)]$DirectoryObject,
        [Parameter(Mandatory)][hashtable]$Cache
    )

    $id = $DirectoryObject.Id

    if ($Cache.ContainsKey($id)) { return $Cache[$id] }

    $props = $DirectoryObject.AdditionalProperties

    $displayName  = $props['displayName']
    $mailNickname = $props['mailNickname']
    $groupTypes   = @($props['groupTypes'])
    $mailEnabled  = [bool]$props['mailEnabled']
    $securityOn   = [bool]$props['securityEnabled']

    # memberOf does not always expand every property; fall back to a direct read.
    if ($null -eq $props['securityEnabled']) {
        try {
            $group = Get-MgGroup -GroupId $id -Property Id,DisplayName,MailNickname,GroupTypes,MailEnabled,SecurityEnabled -ErrorAction Stop

            $displayName  = $group.DisplayName
            $mailNickname = $group.MailNickname
            $groupTypes   = @($group.GroupTypes)
            $mailEnabled  = [bool]$group.MailEnabled
            $securityOn   = [bool]$group.SecurityEnabled
        }
        catch {
            Write-Warning "Could not read group $id : $($_.Exception.Message)"
        }
    }

    $isDynamic = $groupTypes -contains 'DynamicMembership'
    $isUnified = $groupTypes -contains 'Unified'

    if ($isUnified)                    { $type = 'Microsoft 365 Group' }
    elseif ($mailEnabled -and $securityOn) { $type = 'Mail-enabled security group' }
    elseif ($mailEnabled)              { $type = 'Distribution list' }
    elseif ($securityOn)               { $type = 'Security group' }
    else                               { $type = 'Unknown' }

    $importable = $true
    $skipReason = ''

    if ($isDynamic) {
        $importable = $false
        $skipReason = 'Dynamic group - membership is set by a rule, not by adding members'
    }
    elseif ($mailEnabled -and -not $isUnified) {
        $importable = $false
        $skipReason = 'Distribution / mail-enabled security group - requires Exchange Online PowerShell'
    }
    elseif ($type -eq 'Unknown') {
        $importable = $false
        $skipReason = 'Group type could not be determined'
    }

    $facts = [PSCustomObject]@{
        DisplayName  = if ($displayName) { $displayName } else { $id }
        MailNickname = $mailNickname
        GroupType    = $type
        Importable   = $importable
        SkipReason   = $skipReason
    }

    $Cache[$id] = $facts
    return $facts
}

function New-MembershipRow {
    param(
        [Parameter(Mandatory)]$Guest,
        [Parameter(Mandatory)][string]$ExternalEmail,
        $GroupId        = '',
        $GroupName      = '',
        $MailNickname   = '',
        $GroupType      = '',
        $MembershipType = '',
        $Importable     = $false,
        $SkipReason     = ''
    )

    [PSCustomObject]@{
        GuestDisplayName        = $Guest.DisplayName
        ExternalEmail           = $ExternalEmail
        SourceUserPrincipalName = $Guest.UserPrincipalName
        SourceObjectId          = $Guest.Id
        ExternalUserState       = $Guest.ExternalUserState
        AccountEnabled          = $Guest.AccountEnabled
        CompanyName             = $Guest.CompanyName
        JobTitle                = $Guest.JobTitle
        Department              = $Guest.Department
        GroupDisplayName        = $GroupName
        GroupMailNickname       = $MailNickname
        SourceGroupId           = $GroupId
        GroupType               = $GroupType
        MembershipType          = $MembershipType
        Importable              = $Importable
        SkipReason              = $SkipReason
    }
}

# Shared CSV handling - see Common/InputCsv.ps1
$commonPath = Join-Path -Path $PSScriptRoot -ChildPath '..\Common\InputCsv.ps1'

if (-not (Test-Path -Path $commonPath)) {
    throw "Could not find $commonPath. Run this script from inside the complete toolkit folder - it depends on the shared helper in Common/."
}

. $commonPath

#endregion Helpers ------------------------------------------------------------

Initialize-OutputPath -Path $OutputPath
Initialize-OutputPath -Path $UnsupportedOutputPath

Write-Host 'Connecting to Microsoft Graph (source tenant)...' -ForegroundColor Cyan

$connectParams = @{
    Scopes = @('User.Read.All', 'Group.Read.All', 'Directory.Read.All')
}

if ($TenantId) { $connectParams['TenantId'] = $TenantId }

Connect-MgGraph @connectParams

$context = Get-MgContext

Write-Host "Connected to tenant $($context.TenantId) as $($context.Account)" -ForegroundColor Green

Write-Host 'Retrieving guest users...' -ForegroundColor Cyan

$guests = @(
    Get-MgUser -All -Filter "userType eq 'Guest'" -ConsistencyLevel eventual -CountVariable guestCount -Property @(
        'Id', 'DisplayName', 'UserPrincipalName', 'Mail', 'OtherMails', 'AccountEnabled',
        'UserType', 'ExternalUserState', 'CompanyName', 'JobTitle', 'Department', 'CreatedDateTime'
    )
)

if ($guests.Count -eq 0) {
    Write-Host 'No guest users found in this tenant.' -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    return
}

Write-Host "Found $($guests.Count) guest users." -ForegroundColor Green

# Apply the account-state filters.
$skippedDisabled = 0
$skippedPending  = 0

$guests = @(
    $guests | Where-Object {
        if (-not $IncludeDisabledGuests -and -not $_.AccountEnabled) {
            $script:skippedDisabled++
            return $false
        }

        if (-not $IncludePendingAcceptance -and $_.ExternalUserState -eq 'PendingAcceptance') {
            $script:skippedPending++
            return $false
        }

        return $true
    }
)

if ($skippedDisabled) { Write-Host "  Skipping $skippedDisabled disabled guest(s). Use -IncludeDisabledGuests to keep them." -ForegroundColor Yellow }
if ($skippedPending)  { Write-Host "  Skipping $skippedPending guest(s) who never accepted their invitation. Use -IncludePendingAcceptance to keep them." -ForegroundColor Yellow }

if ($guests.Count -eq 0) {
    Write-Host 'No guests remain after filtering. Nothing to export.' -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    return
}

Write-Host "Exporting memberships for $($guests.Count) guest(s)..." -ForegroundColor Cyan

$rows       = [System.Collections.Generic.List[PSCustomObject]]::new()
$groupCache = @{}
$noEmail    = [System.Collections.Generic.List[string]]::new()

$counter = 0

foreach ($guest in $guests) {

    $counter++

    Write-Progress -Activity 'Exporting guest memberships' `
                   -Status "$counter of $($guests.Count) - $($guest.DisplayName)" `
                   -PercentComplete (($counter / $guests.Count) * 100)

    $externalEmail = Resolve-ExternalEmail -User $guest

    if (-not $externalEmail) {
        # Without an external address the guest cannot be re-invited in tenant B.
        $noEmail.Add("$($guest.DisplayName) [$($guest.UserPrincipalName)]")
        continue
    }

    $memberships = @()
    $ownerships  = @()

    try {
        $memberships = @(Get-MgUserMemberOf -UserId $guest.Id -All -ErrorAction Stop |
            Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.group' })
    }
    catch {
        Write-Warning "Could not read memberships for $($guest.UserPrincipalName): $($_.Exception.Message)"
    }

    if (-not $ExcludeGroupOwnership) {
        try {
            $ownerships = @(Get-MgUserOwnedObject -UserId $guest.Id -All -ErrorAction Stop |
                Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.group' })
        }
        catch {
            Write-Warning "Could not read owned groups for $($guest.UserPrincipalName): $($_.Exception.Message)"
        }
    }

    $ownedIds = @{}
    foreach ($owned in $ownerships) { $ownedIds[$owned.Id] = $true }

    if ($memberships.Count -eq 0 -and $ownerships.Count -eq 0) {
        # Still export the guest so the import re-invites them with no group work.
        $rows.Add((New-MembershipRow -Guest $guest -ExternalEmail $externalEmail `
                                     -GroupName 'None' -MembershipType 'None' -Importable $true))
        continue
    }

    foreach ($membership in $memberships) {

        $facts = Get-GroupFacts -DirectoryObject $membership -Cache $groupCache

        # A guest who is both member and owner is recorded as Owner; the import adds
        # owners as members too, so a single row keeps the CSV unambiguous.
        $membershipType = if ($ownedIds.ContainsKey($membership.Id)) { 'Owner' } else { 'Member' }

        $rows.Add((New-MembershipRow -Guest $guest -ExternalEmail $externalEmail `
                                     -GroupId $membership.Id `
                                     -GroupName $facts.DisplayName `
                                     -MailNickname $facts.MailNickname `
                                     -GroupType $facts.GroupType `
                                     -MembershipType $membershipType `
                                     -Importable $facts.Importable `
                                     -SkipReason $facts.SkipReason))
    }

    # Groups the guest owns without being a member of.
    foreach ($owned in $ownerships) {

        if ($memberships.Id -contains $owned.Id) { continue }

        $facts = Get-GroupFacts -DirectoryObject $owned -Cache $groupCache

        $rows.Add((New-MembershipRow -Guest $guest -ExternalEmail $externalEmail `
                                     -GroupId $owned.Id `
                                     -GroupName $facts.DisplayName `
                                     -MailNickname $facts.MailNickname `
                                     -GroupType $facts.GroupType `
                                     -MembershipType 'OwnerOnly' `
                                     -Importable $facts.Importable `
                                     -SkipReason $facts.SkipReason))
    }
}

Write-Progress -Activity 'Exporting guest memberships' -Completed

if ($noEmail.Count -gt 0) {
    Write-Warning "$($noEmail.Count) guest(s) have no resolvable external email address and were skipped:"
    $noEmail | ForEach-Object { Write-Warning "  $_" }
}

$sorted = $rows | Sort-Object GuestDisplayName, GroupDisplayName

$sorted | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

# Confirm the file we just wrote can be read back. A report that cannot be parsed
# is worse than no report, and catching it here beats discovering it at import.
$verified = Assert-WrittenCsv -Path $OutputPath -ExpectedRows $sorted.Count -RequiredColumns @(
    'ExternalEmail', 'GuestDisplayName', 'GroupDisplayName', 'MembershipType', 'Importable'
)

$unsupported = @($sorted | Where-Object { -not $_.Importable -and $_.GroupDisplayName -ne 'None' })

if ($unsupported.Count -gt 0) {
    $unsupported | Export-Csv -Path $UnsupportedOutputPath -NoTypeInformation -Encoding UTF8
}

$uniqueGuests = ($sorted | Select-Object -ExpandProperty ExternalEmail -Unique).Count
$importable   = @($sorted | Where-Object { $_.Importable -and $_.GroupDisplayName -ne 'None' }).Count

Write-Host ''
Write-Host '--- Export summary ---' -ForegroundColor Green
Write-Host "  Guests exported          : $uniqueGuests"
Write-Host "  Membership rows          : $($sorted.Count)"
Write-Host "  Importable via Graph     : $importable"
Write-Host "  Needs manual handling    : $($unsupported.Count)"
Write-Host ''
if ($verified) {
    Write-Host "  Report : $((Resolve-Path -Path $OutputPath).Path)" -ForegroundColor Green
}
else {
    Write-Host "  Report : $((Resolve-Path -Path $OutputPath).Path)  (FAILED VERIFICATION - see warnings above)" -ForegroundColor Red
}

if ($unsupported.Count -gt 0) {
    Write-Host "  Manual : $((Resolve-Path -Path $UnsupportedOutputPath).Path)" -ForegroundColor Yellow
}

Disconnect-MgGraph | Out-Null

Write-Host ''
Write-Host 'Export complete.' -ForegroundColor Green
