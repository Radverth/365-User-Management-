#Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Identity.SignIns

<#
.SYNOPSIS
    Recreates guests and their group memberships in the target tenant from the
    CSV produced by Export-GuestPermissions.ps1.

.DESCRIPTION
    Run this against tenant B, after the groups have been migrated with the same
    display names.

    The script is additive and safe to re-run. It never removes a user, a member
    or an owner, and it never re-invites or re-adds anything that is already
    present:
      - A guest whose external address already exists is reused, not re-invited.
      - Existing group members and owners are detected first and left untouched.
      - Anything that cannot be actioned is logged rather than forced.

    Guests are matched across tenants by their external (home tenant) email
    address. A guest's UPN is tenant-specific and cannot be used for matching.

.PARAMETER InputPath
    The CSV produced by Export-GuestPermissions.ps1.

.PARAMETER SendInvitationMessage
    Send Microsoft's invitation email to each newly invited guest. Omitted by
    default, so guests are created silently and receive nothing.

.PARAMETER CustomInvitationMessage
    Text added to the invitation email. Only has an effect alongside
    -SendInvitationMessage.

.PARAMETER InviteRedirectUrl
    Where the guest lands after redeeming. Defaults to the My Apps portal.

.PARAMETER LogPath
    CSV recording the outcome of every action taken.

.PARAMETER TenantId
    Optional tenant ID or domain to sign in against.

.PARAMETER ResendInvitations
    Also email guests who already exist, not just ones created on this run.

    Use this when guests were added to the tenant by hand, or by an earlier silent
    run, and now need telling. Only has an effect alongside
    -SendInvitationMessage.

    Re-inviting an existing guest does not create a second account and does not
    disturb their group memberships: the invitation is matched on the email
    address, the existing object is reused, and only a fresh invitation email is
    sent.

.PARAMETER SkipInvitations
    Do not create missing guests; only process group membership for guests that
    already exist.

.PARAMETER SkipGroupMembership
    Only create the guests; do not touch group membership at all.

.PARAMETER SkipOwnership
    Add guests as group members but never as group owners.

.PARAMETER Delimiter
    Field separator of the input CSV. Detected automatically when omitted, which
    covers files re-saved by Excel in a locale that uses semicolons.

.PARAMETER MaxGuests
    Process at most this many guests. Useful for a pilot run against a handful of
    accounts before committing to the full migration.

.EXAMPLE
    .\Import-GuestPermissions.ps1 -InputPath .\TenantA_GuestPermissions.csv -WhatIf

    Full dry run. Shows every invitation and membership that would be created.

.EXAMPLE
    .\Import-GuestPermissions.ps1 -InputPath .\TenantA_GuestPermissions.csv -MaxGuests 5

    Pilot: create the first five guests silently, with their group memberships.

.EXAMPLE
    .\Import-GuestPermissions.ps1 -InputPath .\TenantA_GuestPermissions.csv -SendInvitationMessage

    Full run, emailing each new guest their invitation.

.EXAMPLE
    .\Import-GuestPermissions.ps1 -InputPath .\TenantA_GuestPermissions.csv `
        -SendInvitationMessage -ResendInvitations -SkipGroupMembership

    Email everyone in the file, including guests already in the tenant, and change
    no memberships at all. This is how you tell guests you added by hand.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$InputPath,

    [switch]$SendInvitationMessage,

    [string]$CustomInvitationMessage,

    [string]$InviteRedirectUrl = 'https://myapplications.microsoft.com',

    [string]$LogPath = ".\TenantB_GuestImport_Log.csv",

    [string]$TenantId,

    [switch]$ResendInvitations,

    [switch]$SkipInvitations,

    [switch]$SkipGroupMembership,

    [switch]$SkipOwnership,

    [int]$MaxGuests = 0,

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

function ConvertTo-Bool {
    <#  Import-Csv returns every field as a string, so $row.Importable is the text
        'True' rather than a boolean. #>
    param($Value)

    if ($Value -is [bool]) { return $Value }
    if ($null -eq $Value)  { return $false }

    return ([string]$Value).Trim() -in @('True', 'true', 'TRUE', '1', 'Yes')
}

function Resolve-ExternalEmail {
    <#  Same decoding as the export, so a guest already present in the target tenant
        is recognised by the address they were originally invited with. #>
    param([Parameter(Mandatory)]$User)

    if ($User.Mail) { return $User.Mail.Trim() }

    if ($User.OtherMails -and $User.OtherMails.Count -gt 0) {
        return ($User.OtherMails | Select-Object -First 1).Trim()
    }

    $upn = $User.UserPrincipalName

    if ($upn -and $upn -match '^(?<local>.+)#EXT#@') {
        $local = $Matches['local']
        $split = $local.LastIndexOf('_')

        if ($split -gt 0) {
            return ($local.Substring(0, $split) + '@' + $local.Substring($split + 1))
        }
    }

    return $null
}

# Shared CSV input handling - see Common/InputCsv.ps1
$commonPath = Join-Path -Path $PSScriptRoot -ChildPath '..\Common\InputCsv.ps1'

if (-not (Test-Path -Path $commonPath)) {
    throw "Could not find $commonPath. Run this script from inside the complete toolkit folder - it depends on the shared helper in Common/."
}

. $commonPath

#endregion Helpers ------------------------------------------------------------

if (-not (Test-Path -Path $InputPath)) {
    throw "Input file not found: $InputPath"
}

Initialize-OutputPath -Path $LogPath

$rows = Import-InputCsv -Path $InputPath -Delimiter $Delimiter -Expected 'the CSV produced by Export-GuestPermissions.ps1' -RequiredColumns @(
    'ExternalEmail', 'GuestDisplayName', 'GroupDisplayName', 'MembershipType', 'Importable'
)

if ($rows.Count -eq 0) {
    Write-Host 'The input CSV contains no rows. Nothing to do.' -ForegroundColor Yellow
    return
}

Write-Host 'Connecting to Microsoft Graph (target tenant)...' -ForegroundColor Cyan

$connectParams = @{
    Scopes = @('User.Invite.All', 'User.ReadWrite.All', 'Group.ReadWrite.All', 'Directory.ReadWrite.All')
}

if ($TenantId) { $connectParams['TenantId'] = $TenantId }

Connect-MgGraph @connectParams

$context = Get-MgContext

Write-Host "Connected to tenant $($context.TenantId) as $($context.Account)" -ForegroundColor Green
Write-Host ''

if ($SendInvitationMessage) {

    if ($ResendInvitations) {
        Write-Host 'Invitation emails WILL be sent - to new guests AND to guests who already exist.' -ForegroundColor Yellow
        Write-Host 'Existing guests keep their object and their group memberships; only an email is sent.' -ForegroundColor DarkGray
    }
    else {
        Write-Host 'Invitation emails WILL be sent to newly created guests.' -ForegroundColor Yellow
        Write-Host 'Guests who already exist are not emailed. Add -ResendInvitations to include them.' -ForegroundColor DarkGray
    }
}
else {
    Write-Host 'Invitation emails will NOT be sent. Guests are created silently.' -ForegroundColor Cyan

    if ($ResendInvitations) {
        Write-Warning '-ResendInvitations does nothing without -SendInvitationMessage, so no emails will be sent.'
    }
}

Write-Host ''

#region Build target-tenant indexes -------------------------------------------

Write-Host 'Indexing existing guests in the target tenant...' -ForegroundColor Cyan

# One bulk read instead of a filtered query per guest. Every address a guest is
# known by maps to their object ID, so a re-run recognises everyone immediately.
$guestIndex = @{}

$existingGuests = @(
    Get-MgUser -All -Filter "userType eq 'Guest'" -ConsistencyLevel eventual -CountVariable existingGuestCount -Property @(
        'Id', 'DisplayName', 'UserPrincipalName', 'Mail', 'OtherMails', 'UserType'
    )
)

foreach ($guest in $existingGuests) {

    $keys = [System.Collections.Generic.List[string]]::new()

    if ($guest.Mail) { $keys.Add($guest.Mail) }

    foreach ($other in @($guest.OtherMails)) {
        if ($other) { $keys.Add($other) }
    }

    $decoded = Resolve-ExternalEmail -User $guest
    if ($decoded) { $keys.Add($decoded) }

    foreach ($key in $keys) {
        $normalised = $key.Trim().ToLowerInvariant()
        if ($normalised -and -not $guestIndex.ContainsKey($normalised)) {
            $guestIndex[$normalised] = $guest.Id
        }
    }
}

Write-Host "  $($existingGuests.Count) existing guest(s) indexed." -ForegroundColor Green

Write-Host 'Indexing groups in the target tenant...' -ForegroundColor Cyan

# Groups are matched by display name because object IDs differ between tenants.
$groupIndex     = @{}
$ambiguousNames = @{}

$targetGroups = @(
    Get-MgGroup -All -Property Id,DisplayName,MailNickname,GroupTypes,MailEnabled,SecurityEnabled
)

foreach ($group in $targetGroups) {

    if (-not $group.DisplayName) { continue }

    $key = $group.DisplayName.Trim().ToLowerInvariant()

    if ($groupIndex.ContainsKey($key)) {
        # Two groups share a display name - we cannot know which one was intended.
        $ambiguousNames[$key] = $true
    }
    else {
        $groupIndex[$key] = $group
    }
}

Write-Host "  $($targetGroups.Count) group(s) indexed." -ForegroundColor Green

if ($ambiguousNames.Count -gt 0) {
    Write-Warning "$($ambiguousNames.Count) group name(s) are used by more than one group. Memberships for those will be skipped and logged."
}

Write-Host ''

#endregion Build target-tenant indexes ----------------------------------------

# Lazily-filled caches of current members / owners, keyed by target group ID.
$memberCache = @{}
$ownerCache  = @{}

function Get-GroupMemberSet {
    param([Parameter(Mandatory)][string]$GroupId)

    if (-not $memberCache.ContainsKey($GroupId)) {

        $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

        foreach ($member in @(Get-MgGroupMember -GroupId $GroupId -All -ErrorAction Stop)) {
            [void]$set.Add($member.Id)
        }

        $memberCache[$GroupId] = $set
    }

    # The leading comma stops PowerShell unrolling the HashSet into the pipeline,
    # which would return its elements (or nothing at all, when it is empty).
    return ,$memberCache[$GroupId]
}

function Get-GroupOwnerSet {
    param([Parameter(Mandatory)][string]$GroupId)

    if (-not $ownerCache.ContainsKey($GroupId)) {

        $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

        foreach ($owner in @(Get-MgGroupOwner -GroupId $GroupId -All -ErrorAction Stop)) {
            [void]$set.Add($owner.Id)
        }

        $ownerCache[$GroupId] = $set
    }

    return ,$ownerCache[$GroupId]
}

$log = [System.Collections.Generic.List[PSCustomObject]]::new()

function Write-Result {
    param(
        [string]$ExternalEmail,
        [string]$DisplayName,
        [string]$GroupName,
        [string]$Action,
        [string]$Status,
        [string]$Detail = ''
    )

    $log.Add([PSCustomObject]@{
        Timestamp     = (Get-Date).ToString('s')
        ExternalEmail = $ExternalEmail
        DisplayName   = $DisplayName
        GroupName     = $GroupName
        Action        = $Action
        Status        = $Status
        Detail        = $Detail
    })
}

# Group the flat CSV back into one entry per guest.
$guestGroups = $rows | Group-Object -Property ExternalEmail

if ($MaxGuests -gt 0 -and $guestGroups.Count -gt $MaxGuests) {
    Write-Host "Limiting this run to the first $MaxGuests guest(s) of $($guestGroups.Count)." -ForegroundColor Yellow
    $guestGroups = $guestGroups | Select-Object -First $MaxGuests
}

Write-Host "Processing $($guestGroups.Count) guest(s)..." -ForegroundColor Cyan
Write-Host ''

$counter = 0

foreach ($entry in $guestGroups) {

    $counter++

    $externalEmail = $entry.Name
    $guestRows     = @($entry.Group)
    $displayName   = ($guestRows | Select-Object -First 1).GuestDisplayName

    Write-Progress -Activity 'Importing guests' `
                   -Status "$counter of $($guestGroups.Count) - $displayName" `
                   -PercentComplete (($counter / $guestGroups.Count) * 100)

    if ([string]::IsNullOrWhiteSpace($externalEmail)) {
        Write-Result -DisplayName $displayName -Action 'Invite' -Status 'Skipped' `
                     -Detail 'Row has no ExternalEmail value'
        continue
    }

    $lookupKey = $externalEmail.Trim().ToLowerInvariant()
    $userId    = $null

    # True when -WhatIf suppressed the invitation. The guest has no object yet, so
    # membership cannot be applied, but every group name is still resolved against
    # the target tenant - that check is the most useful part of a rehearsal.
    $simulated = $false

    #region Ensure the guest exists -------------------------------------------

    if ($guestIndex.ContainsKey($lookupKey)) {

        $userId = $guestIndex[$lookupKey]

        if ($ResendInvitations -and $SendInvitationMessage) {

            if ($PSCmdlet.ShouldProcess("guest $externalEmail", 'Resend invitation email')) {

                try {
                    $resendParams = @{
                        InvitedUserEmailAddress = $externalEmail
                        InviteRedirectUrl       = $InviteRedirectUrl
                        SendInvitationMessage   = $true
                        ErrorAction             = 'Stop'
                    }

                    if ($displayName) { $resendParams['InvitedUserDisplayName'] = $displayName }

                    if ($CustomInvitationMessage) {
                        $resendParams['InvitedUserMessageInfo'] = @{
                            CustomizedMessageBody = $CustomInvitationMessage
                        }
                    }

                    # Matched on the email address, so this reuses the existing
                    # object rather than creating a second one, and leaves group
                    # membership alone. Only a fresh invitation email is sent.
                    $resent = New-MgInvitation @resendParams

                    $resentId = $resent.InvitedUser.Id

                    if ($resentId -and $resentId -ne $userId) {
                        # Should not happen, but if the directory returned a
                        # different object the operator needs to know before it
                        # gets memberships added to it.
                        Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -Action 'Invite' `
                                     -Status 'Failed' `
                                     -Detail "Resend returned object $resentId but $userId was expected. Check for a duplicate guest before continuing."

                        Write-Warning "Resending to $externalEmail returned a different object; skipping group work for this guest."
                        continue
                    }

                    Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -Action 'Invite' `
                                 -Status 'InvitationResent' -Detail "Existing object $userId reused; invitation email sent"
                }
                catch {
                    Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -Action 'Invite' `
                                 -Status 'Failed' -Detail "Resend failed: $($_.Exception.Message)"

                    Write-Warning "Could not resend to $externalEmail : $($_.Exception.Message)"
                }
            }
            else {
                Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -Action 'Invite' `
                             -Status 'WhatIf' -Detail "Would resend the invitation email to existing object $userId"
            }
        }
        else {
            Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -Action 'Invite' `
                         -Status 'AlreadyExists' -Detail "Existing object $userId reused"
        }
    }
    elseif ($SkipInvitations) {

        Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -Action 'Invite' `
                     -Status 'Skipped' -Detail '-SkipInvitations was specified and the guest does not exist'
        continue
    }
    else {

        $target = "guest $externalEmail"
        $verb   = if ($SendInvitationMessage) { 'Invite (sending email)' } else { 'Invite (no email)' }

        if ($PSCmdlet.ShouldProcess($target, $verb)) {

            try {
                $invitationParams = @{
                    InvitedUserEmailAddress = $externalEmail
                    InviteRedirectUrl       = $InviteRedirectUrl
                    SendInvitationMessage   = [bool]$SendInvitationMessage
                    ErrorAction             = 'Stop'
                }

                if ($displayName) { $invitationParams['InvitedUserDisplayName'] = $displayName }

                if ($SendInvitationMessage -and $CustomInvitationMessage) {
                    $invitationParams['InvitedUserMessageInfo'] = @{
                        CustomizedMessageBody = $CustomInvitationMessage
                    }
                }

                $invitation = New-MgInvitation @invitationParams

                $userId = $invitation.InvitedUser.Id

                # Keep the index current so a duplicate row later in the file is a hit.
                $guestIndex[$lookupKey] = $userId

                Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -Action 'Invite' `
                             -Status 'Created' -Detail "New object $userId; email sent: $([bool]$SendInvitationMessage)"
            }
            catch {
                Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -Action 'Invite' `
                             -Status 'Failed' -Detail $_.Exception.Message

                Write-Warning "Could not invite $externalEmail : $($_.Exception.Message)"
                continue
            }
        }
        else {
            Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -Action 'Invite' `
                         -Status 'WhatIf' -Detail "Would invite; email would be sent: $([bool]$SendInvitationMessage)"

            $simulated = $true
        }
    }

    #endregion Ensure the guest exists ----------------------------------------

    if ($SkipGroupMembership) { continue }

    #region Group membership --------------------------------------------------

    foreach ($row in $guestRows) {

        $groupName = $row.GroupDisplayName

        if ([string]::IsNullOrWhiteSpace($groupName) -or $groupName -eq 'None') { continue }

        if (-not (ConvertTo-Bool $row.Importable)) {
            Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -GroupName $groupName `
                         -Action 'AddMember' -Status 'Skipped' `
                         -Detail $(if ($row.SkipReason) { $row.SkipReason } else { 'Marked not importable by the export' })
            continue
        }

        $groupKey = $groupName.Trim().ToLowerInvariant()

        if ($ambiguousNames.ContainsKey($groupKey)) {
            Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -GroupName $groupName `
                         -Action 'AddMember' -Status 'Skipped' `
                         -Detail 'More than one group in the target tenant uses this display name'
            continue
        }

        if (-not $groupIndex.ContainsKey($groupKey)) {
            Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -GroupName $groupName `
                         -Action 'AddMember' -Status 'GroupNotFound' `
                         -Detail 'No group with this display name exists in the target tenant'
            continue
        }

        $targetGroup = $groupIndex[$groupKey]

        if (@($targetGroup.GroupTypes) -contains 'DynamicMembership') {
            Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -GroupName $groupName `
                         -Action 'AddMember' -Status 'Skipped' `
                         -Detail 'Target group uses dynamic membership; members cannot be added directly'
            continue
        }

        # --- member ---------------------------------------------------------

        if ($simulated) {
            Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -GroupName $groupName `
                         -Action 'AddMember' -Status 'WhatIf' `
                         -Detail 'Group resolved in the target tenant; would add as member once the guest exists'

            if (-not $SkipOwnership -and $row.MembershipType -in @('Owner', 'OwnerOnly')) {
                Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -GroupName $groupName `
                             -Action 'AddOwner' -Status 'WhatIf' `
                             -Detail 'Group resolved in the target tenant; would add as owner once the guest exists'
            }

            continue
        }

        try {
            $members = Get-GroupMemberSet -GroupId $targetGroup.Id
        }
        catch {
            Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -GroupName $groupName `
                         -Action 'AddMember' -Status 'Failed' -Detail "Could not read members: $($_.Exception.Message)"
            continue
        }

        if ($members.Contains($userId)) {
            Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -GroupName $groupName `
                         -Action 'AddMember' -Status 'AlreadyMember'
        }
        elseif ($PSCmdlet.ShouldProcess("$groupName", "Add $externalEmail as member")) {

            try {
                New-MgGroupMember -GroupId $targetGroup.Id -DirectoryObjectId $userId -ErrorAction Stop

                [void]$members.Add($userId)

                Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -GroupName $groupName `
                             -Action 'AddMember' -Status 'Added'
            }
            catch {
                Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -GroupName $groupName `
                             -Action 'AddMember' -Status 'Failed' -Detail $_.Exception.Message
            }
        }
        else {
            Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -GroupName $groupName `
                         -Action 'AddMember' -Status 'WhatIf' -Detail 'Would add as member'
        }

        # --- owner ----------------------------------------------------------

        if ($SkipOwnership) { continue }
        if ($row.MembershipType -notin @('Owner', 'OwnerOnly')) { continue }

        try {
            $owners = Get-GroupOwnerSet -GroupId $targetGroup.Id
        }
        catch {
            Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -GroupName $groupName `
                         -Action 'AddOwner' -Status 'Failed' -Detail "Could not read owners: $($_.Exception.Message)"
            continue
        }

        if ($owners.Contains($userId)) {
            Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -GroupName $groupName `
                         -Action 'AddOwner' -Status 'AlreadyOwner'
        }
        elseif ($PSCmdlet.ShouldProcess("$groupName", "Add $externalEmail as owner")) {

            try {
                New-MgGroupOwner -GroupId $targetGroup.Id -DirectoryObjectId $userId -ErrorAction Stop

                [void]$owners.Add($userId)

                Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -GroupName $groupName `
                             -Action 'AddOwner' -Status 'Added'
            }
            catch {
                Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -GroupName $groupName `
                             -Action 'AddOwner' -Status 'Failed' -Detail $_.Exception.Message
            }
        }
        else {
            Write-Result -ExternalEmail $externalEmail -DisplayName $displayName -GroupName $groupName `
                         -Action 'AddOwner' -Status 'WhatIf' -Detail 'Would add as owner'
        }
    }

    #endregion Group membership -----------------------------------------------
}

Write-Progress -Activity 'Importing guests' -Completed

# -WhatIf:$false so a rehearsal still produces its report - the log is the whole
# point of a dry run.
$log | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8 -WhatIf:$false

Write-Host ''
Write-Host '--- Import summary ---' -ForegroundColor Green

$log | Group-Object Action, Status | Sort-Object Name | ForEach-Object {
    Write-Host ("  {0,-32} {1}" -f $_.Name, $_.Count)
}

$failures = @($log | Where-Object { $_.Status -eq 'Failed' })

Write-Host ''
Write-Host "  Log : $((Resolve-Path -Path $LogPath).Path)" -ForegroundColor Green

if ($failures.Count -gt 0) {
    Write-Host "  $($failures.Count) action(s) failed - see the log for details." -ForegroundColor Red
}

Disconnect-MgGraph | Out-Null

Write-Host ''
Write-Host 'Import complete. Re-running is safe: nothing already present is changed.' -ForegroundColor Green
