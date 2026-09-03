#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Rewrites mailbox email addresses from a spreadsheet, for the two moves a
    cross-tenant migration needs: releasing a vanity address in the old tenant,
    and adding it as an alias in the new one.

.DESCRIPTION
    Two jobs, one spreadsheet, chosen with -Action.

    SetPrimaryToOnMicrosoft  (run against the OLD tenant)
        Finds each mailbox by its PrimarySMTPAddress and makes the .onmicrosoft
        address in the Email Address column the primary instead. The old primary
        is kept as an alias unless -RemoveOldPrimary is given, which is what
        actually frees the domain for verification in the new tenant.

    AddAliasToNewPrimary     (run against the NEW tenant)
        Finds each mailbox by the address in the Hawsons Primary column and adds
        the PrimarySMTPAddress to it as an alias. The primary address is left
        exactly as it is, so nobody's reply-to or sending address changes.

    Both are additive where they can be and safe to re-run: a mailbox that is
    already in the wanted state is reported and skipped, not written to. Nothing
    outside the SMTP addresses is touched - SIP, X500 and EUM entries are carried
    across untouched, so free/busy and old Outlook replies keep working.

    Every mailbox is checked before it is written to. Directory-synced mailboxes,
    addresses on a domain the tenant has not accepted, and addresses already used
    by somebody else are reported with the reason rather than attempted.

.PARAMETER InputPath
    Spreadsheet of mailboxes, saved as CSV. Column names are matched loosely, so
    'PrimarySMTPAddress', 'Primary SMTP Address' and 'primary smtp address' are
    all understood. These are the columns it looks for:

        User Principal Name   optional; used to find the mailbox if the address
                              lookup fails, and shown in the log
        Hawsons Primary       the address that is to stay primary in the new
                              tenant. Required by AddAliasToNewPrimary
        PrimarySMTPAddress    the current primary in the old tenant. Required by
                              both actions
        Email Address         the .onmicrosoft address that is to become primary.
                              Required by SetPrimaryToOnMicrosoft

.PARAMETER Action
    SetPrimaryToOnMicrosoft or AddAliasToNewPrimary. See the description.

.PARAMETER RemoveOldPrimary
    SetPrimaryToOnMicrosoft only. Also remove the old primary address from the
    mailbox instead of leaving it behind as an alias.

    You need this to release the domain: a domain cannot be verified in the new
    tenant while addresses in the old one still use it. It is off by default
    because it is the one step here that loses something - mail sent to the old
    address stops being delivered the moment it is removed.

.PARAMETER AllowAnyNewPrimary
    Accept a new primary address that is not on an .onmicrosoft.com domain.

    Without it, a value in the Email Address column that is not an .onmicrosoft
    address is refused. That check exists because the usual way to get this wrong
    is to point the script at the wrong column, and the damage is done before you
    notice.

.PARAMETER DisableEmailAddressPolicy
    Turn off the email address policy on a mailbox that has one, so its addresses
    can be set. Exchange refuses to change addresses on a policy-managed mailbox,
    and this is the supported way round it. Without this, such mailboxes are
    reported and skipped.

.PARAMETER LogPath
    CSV recording the outcome for every row. Defaults to a name that includes the
    action, so the two runs do not overwrite each other.

.PARAMETER MaxMailboxes
    Stop after this many rows. For a pilot run over a handful of mailboxes before
    committing to the whole file.

.PARAMETER AdminUpn
    Sign in as this account. Prompts if omitted.

.PARAMETER Organization
    Tenant domain, for app-only sign-in. Needs -AppId and -CertificateThumbprint.

.PARAMETER AppId
    Application (client) ID, for app-only sign-in.

.PARAMETER CertificateThumbprint
    Certificate thumbprint, for app-only sign-in.

.PARAMETER Delimiter
    Field separator of the input CSV. Detected automatically when omitted.

.EXAMPLE
    .\Set-MailboxAddresses.ps1 -InputPath .\mailboxes.csv -Action SetPrimaryToOnMicrosoft -WhatIf

    Rehearsal against the old tenant. Reports what each mailbox's addresses would
    become, and every reason a mailbox could not be done, without writing.

.EXAMPLE
    .\Set-MailboxAddresses.ps1 -InputPath .\mailboxes.csv -Action SetPrimaryToOnMicrosoft -MaxMailboxes 5

    Pilot: switch the first five mailboxes to their .onmicrosoft address, keeping
    the old address on each as an alias.

.EXAMPLE
    .\Set-MailboxAddresses.ps1 -InputPath .\mailboxes.csv -Action SetPrimaryToOnMicrosoft -RemoveOldPrimary

    The real release: the .onmicrosoft address becomes primary and the old address
    is removed, so the domain can be verified in the new tenant.

.EXAMPLE
    .\Set-MailboxAddresses.ps1 -InputPath .\mailboxes.csv -Action AddAliasToNewPrimary

    Against the new tenant: add each old address as an alias to the mailbox that
    already holds the Hawsons primary. No primary address changes.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$InputPath,

    [Parameter(Mandatory)]
    [ValidateSet('SetPrimaryToOnMicrosoft', 'AddAliasToNewPrimary')]
    [string]$Action,

    [switch]$RemoveOldPrimary,

    [switch]$AllowAnyNewPrimary,

    [switch]$DisableEmailAddressPolicy,

    [string]$LogPath,

    [int]$MaxMailboxes = 0,

    [string]$AdminUpn,

    [string]$Organization,

    [string]$AppId,

    [string]$CertificateThumbprint,

    [string]$Delimiter
)

$ErrorActionPreference = 'Stop'

if (-not $LogPath) { $LogPath = ".\Mailbox_$($Action)_Log.csv" }

#region Helpers ---------------------------------------------------------------

function Initialize-OutputPath {
    param([Parameter(Mandatory)][string]$Path)

    $parent = Split-Path -Path $Path -Parent

    if ($parent -and -not (Test-Path -Path $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
}

function Get-NormalisedName {
    <#  'Primary SMTP Address', 'PrimarySMTPAddress' and 'primary_smtp_address'
        are the same column as far as this script is concerned. Spreadsheets get
        retyped by hand and the spacing never survives. #>
    param([string]$Name)

    if (-not $Name) { return '' }

    return ($Name -replace '[^A-Za-z0-9]', '').ToLowerInvariant()
}

function Resolve-ColumnName {
    <#  Finds which of the file's headers holds a given field, by normalised name.
        Returns the header exactly as it appears in the file, so rows can be read
        with it. #>
    param(
        [Parameter(Mandatory)][string[]]$Headers,
        [Parameter(Mandatory)][string[]]$Accepts
    )

    foreach ($accept in $Accepts) {
        $wanted = Get-NormalisedName -Name $accept

        foreach ($header in $Headers) {
            if ((Get-NormalisedName -Name $header) -eq $wanted) { return $header }
        }
    }

    return $null
}

function Get-RowValue {
    param($Row, [string]$Column)

    if (-not $Column) { return $null }

    $value = $Row.$Column

    if ($null -eq $value) { return $null }

    $text = ([string]$value).Trim()

    if ([string]::IsNullOrWhiteSpace($text)) { return $null }

    return $text
}

function Get-AddressDomain {
    param([string]$Address)

    if (-not $Address) { return $null }

    $at = $Address.LastIndexOf('@')

    if ($at -lt 0 -or $at -eq $Address.Length - 1) { return $null }

    return $Address.Substring($at + 1).ToLowerInvariant()
}

function Test-LooksLikeAddress {
    param([string]$Address)

    return [bool]($Address -match '^[^@\s]+@[^@\s]+\.[^@\s]+$')
}

function Split-ProxyAddress {
    <#  Proxy addresses arrive as 'SMTP:a@b.com', 'smtp:c@b.com', 'SIP:...' and
        'X500:...'. Case of the prefix is what marks the primary, so it has to be
        preserved rather than normalised away. #>
    param([string]$Proxy)

    $colon = $Proxy.IndexOf(':')

    if ($colon -lt 0) {
        # No prefix at all means SMTP by convention, and an alias rather than the
        # primary - Exchange only ever writes the primary with a prefix.
        return [pscustomobject]@{ Prefix = 'smtp'; Address = $Proxy; IsSmtp = $true; IsPrimary = $false }
    }

    $prefix  = $Proxy.Substring(0, $colon)
    $address = $Proxy.Substring($colon + 1)
    $isSmtp  = $prefix -ieq 'smtp'

    return [pscustomobject]@{
        Prefix    = $prefix
        Address   = $address
        IsSmtp    = $isSmtp
        IsPrimary = ($isSmtp -and $prefix -ceq 'SMTP')
    }
}

#endregion Helpers ------------------------------------------------------------

# Shared CSV input handling - see Common/InputCsv.ps1
$commonPath = Join-Path -Path $PSScriptRoot -ChildPath '..\Common\InputCsv.ps1'

if (-not (Test-Path -Path $commonPath)) {
    throw "Could not find $commonPath. Run this script from inside the complete toolkit folder - it depends on the shared helper in Common/."
}

. $commonPath

Initialize-OutputPath -Path $LogPath

#region Read the spreadsheet --------------------------------------------------

Write-Host "Reading $InputPath ..." -ForegroundColor Cyan

# No required columns are declared here: the headers are matched loosely a moment
# later, and a strict match on names typed by hand would reject good files.
$rows = Import-InputCsv -Path $InputPath -Delimiter $Delimiter `
                        -Expected 'the mailbox spreadsheet (user principal name, Hawsons primary, PrimarySMTPAddress, email address)'

if ($rows.Count -eq 0) { throw "No rows found in $InputPath." }

$headers = @($rows[0].PSObject.Properties.Name)

$columns = [ordered]@{
    UserPrincipalName = Resolve-ColumnName -Headers $headers -Accepts @(
        'UserPrincipalName', 'User Principal Name', 'User Principle Name', 'UPN', 'User', 'Identity'
    )
    NewPrimary        = Resolve-ColumnName -Headers $headers -Accepts @(
        'Hawsons Primary', 'HawsonsPrimary', 'New Primary', 'NewPrimary', 'Target Primary', 'TargetPrimary'
    )
    OldPrimary        = Resolve-ColumnName -Headers $headers -Accepts @(
        'PrimarySMTPAddress', 'Primary SMTP Address', 'PrimarySMTP', 'Primary Address', 'Primary'
    )
    OnMicrosoft       = Resolve-ColumnName -Headers $headers -Accepts @(
        'Email Address', 'EmailAddress', 'OnMicrosoftAddress', 'OnMicrosoft', 'Email'
    )
}

# Each action needs a different pair of columns, so only complain about the ones
# it will actually read.
$required = if ($Action -eq 'SetPrimaryToOnMicrosoft') {
    [ordered]@{ OldPrimary = 'PrimarySMTPAddress'; OnMicrosoft = 'Email Address' }
}
else {
    [ordered]@{ NewPrimary = 'Hawsons Primary'; OldPrimary = 'PrimarySMTPAddress' }
}

$missing = @($required.Keys | Where-Object { -not $columns[$_] })

if ($missing.Count -gt 0) {

    $wanted = @($missing | ForEach-Object { $required[$_] })

    $message = [System.Text.StringBuilder]::new()

    [void]$message.AppendLine("Input CSV is missing the column(s) -Action $Action needs: $($wanted -join ', ').")
    [void]$message.AppendLine("  File   : $InputPath")
    [void]$message.AppendLine("  Found  : $($headers -join ', ')")
    [void]$message.AppendLine('')
    [void]$message.AppendLine('  Spacing and capitals do not matter, but the wording does. Rename the header')
    [void]$message.AppendLine('  in the spreadsheet to one of the names above and save it as CSV again.')

    throw $message.ToString()
}

foreach ($key in $columns.Keys) {
    if ($columns[$key]) { Write-Host "  $key -> column '$($columns[$key])'" -ForegroundColor DarkGray }
}

Write-Host "  $($rows.Count) row(s) read." -ForegroundColor Green
Write-Host ''

#endregion Read the spreadsheet -----------------------------------------------

#region Connect ---------------------------------------------------------------

Write-Host 'Connecting to Exchange Online...' -ForegroundColor Cyan

$connectParams = @{ ShowBanner = $false }

if ($AppId -or $CertificateThumbprint -or $Organization) {

    if (-not ($AppId -and $CertificateThumbprint -and $Organization)) {
        throw 'App-only sign-in needs all three of -AppId, -CertificateThumbprint and -Organization. Omit all three to sign in interactively.'
    }

    $connectParams['AppId']                 = $AppId
    $connectParams['CertificateThumbprint'] = $CertificateThumbprint
    $connectParams['Organization']          = $Organization
}
elseif ($AdminUpn) {
    $connectParams['UserPrincipalName'] = $AdminUpn
}

Connect-ExchangeOnline @connectParams

Write-Host '  Connected.' -ForegroundColor Green

# Loaded once. Every failure this catches would otherwise surface as an opaque
# Exchange error on the row that hit it.
$acceptedDomains = @{}

foreach ($domain in @(Get-AcceptedDomain)) {
    $acceptedDomains[$domain.DomainName.ToString().ToLowerInvariant()] = $true
}

Write-Host "  $($acceptedDomains.Count) accepted domain(s) in this tenant." -ForegroundColor DarkGray
Write-Host ''

#endregion Connect ------------------------------------------------------------

$log = [System.Collections.Generic.List[PSCustomObject]]::new()

function Write-Result {
    param(
        [string]$UserPrincipalName,
        [string]$Mailbox,
        [string]$Address,
        [Parameter(Mandatory)][string]$Status,
        [string]$PrimaryBefore,
        [string]$PrimaryAfter,
        [string]$Detail
    )

    $log.Add([PSCustomObject]@{
        Timestamp         = (Get-Date).ToString('s')
        Action            = $Action
        UserPrincipalName = $UserPrincipalName
        Mailbox           = $Mailbox
        Address           = $Address
        Status            = $Status
        PrimaryBefore     = $PrimaryBefore
        PrimaryAfter      = $PrimaryAfter
        Detail            = $Detail
    })
}

#region Process ---------------------------------------------------------------

$toProcess = if ($MaxMailboxes -gt 0) { @($rows | Select-Object -First $MaxMailboxes) } else { $rows }

if ($MaxMailboxes -gt 0 -and $rows.Count -gt $MaxMailboxes) {
    Write-Host "Pilot run: $MaxMailboxes of $($rows.Count) row(s)." -ForegroundColor Yellow
    Write-Host ''
}

Write-Host "Processing $($toProcess.Count) row(s)..." -ForegroundColor Cyan
Write-Host ''

$rowNumber = 0

foreach ($row in $toProcess) {

    $rowNumber++

    Write-Progress -Activity "Updating mailbox addresses ($Action)" `
                   -Status "$rowNumber of $($toProcess.Count)" `
                   -PercentComplete (($rowNumber / [double]$toProcess.Count) * 100)

    $upn         = Get-RowValue -Row $row -Column $columns['UserPrincipalName']
    $newPrimary  = Get-RowValue -Row $row -Column $columns['NewPrimary']
    $oldPrimary  = Get-RowValue -Row $row -Column $columns['OldPrimary']
    $onMicrosoft = Get-RowValue -Row $row -Column $columns['OnMicrosoft']

    # What this row is looking the mailbox up by, and what address it is about.
    if ($Action -eq 'SetPrimaryToOnMicrosoft') {
        $lookup       = $oldPrimary
        $lookupLabel  = 'PrimarySMTPAddress'
        $subject      = $onMicrosoft
        $subjectLabel = 'Email Address'
    }
    else {
        $lookup       = $newPrimary
        $lookupLabel  = 'Hawsons Primary'
        $subject      = $oldPrimary
        $subjectLabel = 'PrimarySMTPAddress'
    }

    if (-not $lookup -and -not $upn) {
        Write-Result -Status 'Skipped' -Detail "Row $rowNumber has no $lookupLabel and no user principal name, so there is nothing to find"
        continue
    }

    if (-not $subject) {
        Write-Result -UserPrincipalName $upn -Address $lookup -Status 'Skipped' `
                     -Detail "Row $rowNumber has no $subjectLabel value"
        continue
    }

    if (-not (Test-LooksLikeAddress -Address $subject)) {
        Write-Result -UserPrincipalName $upn -Address $subject -Status 'Skipped' `
                     -Detail "'$subject' in $subjectLabel is not an email address"
        continue
    }

    #region Find the mailbox --------------------------------------------------

    $mailbox   = $null
    $identity  = $null
    $triedWith = @()

    foreach ($candidate in @($lookup, $upn)) {

        if (-not $candidate -or $candidate -in $triedWith) { continue }

        $triedWith += $candidate

        try {
            $mailbox = Get-Mailbox -Identity $candidate -ErrorAction Stop
            $identity = $candidate
            break
        }
        catch { $mailbox = $null }
    }

    if (-not $mailbox) {

        # Saying what the address IS beats saying it is not a mailbox. A shared
        # mailbox is fine; a mail user or a group is not, and needs a different fix.
        $whatIsIt = $null

        foreach ($candidate in $triedWith) {
            try {
                $recipient = Get-Recipient -Identity $candidate -ErrorAction Stop
                $whatIsIt  = "$candidate exists but is a $($recipient.RecipientTypeDetails), not a mailbox"
                break
            }
            catch { }
        }

        $detail = if ($whatIsIt) { $whatIsIt } else { "No mailbox found for $($triedWith -join ' or ')" }

        Write-Result -UserPrincipalName $upn -Address $lookup -Status 'NotFound' -Detail $detail
        continue
    }

    $mailboxName    = $mailbox.PrimarySmtpAddress.ToString()
    $currentPrimary = $mailboxName
    $writeIdentity  = $mailbox.Guid.ToString()

    #endregion Find the mailbox -----------------------------------------------

    #region Checks that apply to both actions ---------------------------------

    $domain = Get-AddressDomain -Address $subject

    if (-not $acceptedDomains.ContainsKey($domain)) {
        Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'DomainNotAccepted' `
                     -PrimaryBefore $currentPrimary `
                     -Detail "'$domain' is not an accepted domain in this tenant, so the address cannot be added. Add and verify the domain first."
        continue
    }

    $proxies = @($mailbox.EmailAddresses | ForEach-Object { Split-ProxyAddress -Proxy ([string]$_) })
    $existing = @($proxies | Where-Object { $_.IsSmtp } | ForEach-Object { $_.Address })

    #endregion Checks that apply to both actions ------------------------------

    if ($Action -eq 'SetPrimaryToOnMicrosoft') {

        #region Make the .onmicrosoft address primary -------------------------

        if ($currentPrimary -ieq $subject) {

            $stillThere = @($existing | Where-Object { $_ -ieq $oldPrimary }).Count -gt 0

            $detail = if ($RemoveOldPrimary -and $stillThere) {
                "Already primary, but $oldPrimary is still on the mailbox as an alias"
            }
            else {
                'Already the primary address'
            }

            # The one case where "already done" still needs a write: the primary is
            # right but -RemoveOldPrimary was added on a later run.
            if (-not ($RemoveOldPrimary -and $stillThere)) {
                Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'AlreadyPrimary' `
                             -PrimaryBefore $currentPrimary -PrimaryAfter $currentPrimary -Detail $detail
                continue
            }
        }
        elseif ($oldPrimary -and $currentPrimary -ine $oldPrimary) {
            Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'Mismatch' `
                         -PrimaryBefore $currentPrimary `
                         -Detail "The mailbox found has primary $currentPrimary, but the spreadsheet says $oldPrimary. Left alone - check the row."
            continue
        }

        if (-not $AllowAnyNewPrimary -and $domain -notlike '*.onmicrosoft.com') {
            Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'Refused' `
                         -PrimaryBefore $currentPrimary `
                         -Detail "$subject is not an .onmicrosoft.com address. If that is deliberate, re-run with -AllowAnyNewPrimary."
            continue
        }

        if ($mailbox.EmailAddressPolicyEnabled -and -not $DisableEmailAddressPolicy) {
            Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'PolicyManaged' `
                         -PrimaryBefore $currentPrimary `
                         -Detail 'An email address policy manages this mailbox, so Exchange will not let its addresses be set. Re-run with -DisableEmailAddressPolicy.'
            continue
        }

        if ($mailbox.IsDirSynced) {
            Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'DirSynced' `
                         -PrimaryBefore $currentPrimary `
                         -Detail 'This mailbox is synchronised from on-premises Active Directory. Change its proxyAddresses there instead.'
            continue
        }

        # Rebuild the whole list rather than adding to it: only a full array can
        # express "this one is primary and the previous one is not" in one write.
        $newAddresses = [System.Collections.Generic.List[string]]::new()
        $seen         = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        [void]$newAddresses.Add("SMTP:$subject")
        [void]$seen.Add($subject)

        foreach ($proxy in $proxies) {

            if (-not $proxy.IsSmtp) {
                # SIP, X500 and EUM entries are what keep free/busy and replies to
                # old messages working. They are carried across exactly as they are.
                [void]$newAddresses.Add("$($proxy.Prefix):$($proxy.Address)")
                continue
            }

            if ($seen.Contains($proxy.Address)) { continue }

            if ($RemoveOldPrimary -and $oldPrimary -and $proxy.Address -ieq $oldPrimary) { continue }

            [void]$seen.Add($proxy.Address)
            [void]$newAddresses.Add("smtp:$($proxy.Address)")
        }

        $removing = $RemoveOldPrimary -and $oldPrimary -and @($existing | Where-Object { $_ -ieq $oldPrimary }).Count -gt 0

        $description = if ($removing) {
            "Set primary to $subject and remove $oldPrimary"
        }
        else {
            "Set primary to $subject, keeping $currentPrimary as an alias"
        }

        if ($PSCmdlet.ShouldProcess("mailbox $mailboxName", $description)) {

            try {
                $setParams = @{
                    Identity       = $writeIdentity
                    EmailAddresses = $newAddresses.ToArray()
                    ErrorAction    = 'Stop'
                }

                if ($DisableEmailAddressPolicy -and $mailbox.EmailAddressPolicyEnabled) {
                    $setParams['EmailAddressPolicyEnabled'] = $false
                }

                Set-Mailbox @setParams

                # Read it back. Exchange accepts an address list and then applies its
                # own rules to it, so what was asked for is not always what landed.
                $after       = Get-Mailbox -Identity $writeIdentity -ErrorAction Stop
                $afterPrimary = $after.PrimarySmtpAddress.ToString()

                if ($afterPrimary -ieq $subject) {
                    $detail = if ($removing) { "$oldPrimary removed" } else { "$currentPrimary kept as an alias" }

                    Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'PrimaryChanged' `
                                 -PrimaryBefore $currentPrimary -PrimaryAfter $afterPrimary -Detail $detail
                }
                else {
                    Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'Failed' `
                                 -PrimaryBefore $currentPrimary -PrimaryAfter $afterPrimary `
                                 -Detail "Exchange accepted the change but the primary is $afterPrimary, not $subject."

                    Write-Warning "$mailboxName : primary is $afterPrimary, not $subject."
                }
            }
            catch {
                Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'Failed' `
                             -PrimaryBefore $currentPrimary -Detail $_.Exception.Message

                Write-Warning "$mailboxName : $($_.Exception.Message)"
            }
        }
        else {
            Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'WhatIf' `
                         -PrimaryBefore $currentPrimary -PrimaryAfter $subject -Detail "Would $($description.Substring(0,1).ToLowerInvariant())$($description.Substring(1))"
        }

        #endregion Make the .onmicrosoft address primary ----------------------
    }
    else {

        #region Add the old address as an alias -------------------------------

        if ($currentPrimary -ieq $subject) {
            Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'AlreadyPrimary' `
                         -PrimaryBefore $currentPrimary -PrimaryAfter $currentPrimary `
                         -Detail "$subject is already the PRIMARY address on this mailbox, not an alias. Left alone - adding it would change nothing, and demoting it is not this script's job."
            continue
        }

        if (@($existing | Where-Object { $_ -ieq $subject }).Count -gt 0) {
            Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'AlreadyAlias' `
                         -PrimaryBefore $currentPrimary -PrimaryAfter $currentPrimary `
                         -Detail 'Already on the mailbox as an alias'
            continue
        }

        if ($newPrimary -and $currentPrimary -ine $newPrimary) {
            Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'Mismatch' `
                         -PrimaryBefore $currentPrimary `
                         -Detail "The mailbox found has primary $currentPrimary, but the spreadsheet says $newPrimary. Left alone - check the row."
            continue
        }

        # An address already in use elsewhere is the failure that reads worst in
        # Exchange's own words, so it is caught here and named.
        $conflict = $null

        try {
            $holder = Get-Recipient -Identity $subject -ErrorAction Stop

            if ($holder -and $holder.Guid.ToString() -ne $mailbox.Guid.ToString()) {
                $conflict = "$subject is already in use by $($holder.DisplayName) ($($holder.PrimarySmtpAddress)). Remove it there first."
            }
        }
        catch { }

        if ($conflict) {
            Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'Conflict' `
                         -PrimaryBefore $currentPrimary -Detail $conflict
            continue
        }

        if ($mailbox.EmailAddressPolicyEnabled -and -not $DisableEmailAddressPolicy) {
            Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'PolicyManaged' `
                         -PrimaryBefore $currentPrimary `
                         -Detail 'An email address policy manages this mailbox, so Exchange will not let its addresses be set. Re-run with -DisableEmailAddressPolicy.'
            continue
        }

        if ($mailbox.IsDirSynced) {
            Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'DirSynced' `
                         -PrimaryBefore $currentPrimary `
                         -Detail 'This mailbox is synchronised from on-premises Active Directory. Add the alias to its proxyAddresses there instead.'
            continue
        }

        if ($PSCmdlet.ShouldProcess("mailbox $mailboxName", "Add alias $subject, leaving $currentPrimary as the primary")) {

            try {
                $setParams = @{
                    Identity       = $writeIdentity
                    # Lowercase 'smtp:' is what makes it an alias. An uppercase
                    # prefix here would silently take over as the primary.
                    EmailAddresses = @{ Add = "smtp:$subject" }
                    ErrorAction    = 'Stop'
                }

                if ($DisableEmailAddressPolicy -and $mailbox.EmailAddressPolicyEnabled) {
                    $setParams['EmailAddressPolicyEnabled'] = $false
                }

                Set-Mailbox @setParams

                $after        = Get-Mailbox -Identity $writeIdentity -ErrorAction Stop
                $afterPrimary = $after.PrimarySmtpAddress.ToString()
                $afterSmtp    = @($after.EmailAddresses |
                                    ForEach-Object { Split-ProxyAddress -Proxy ([string]$_) } |
                                    Where-Object { $_.IsSmtp } |
                                    ForEach-Object { $_.Address })

                if ($afterPrimary -ine $currentPrimary) {
                    Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'Failed' `
                                 -PrimaryBefore $currentPrimary -PrimaryAfter $afterPrimary `
                                 -Detail "The alias was added but the primary changed from $currentPrimary to $afterPrimary. Put it back before going any further."

                    Write-Warning "$mailboxName : primary changed to $afterPrimary."
                }
                elseif (@($afterSmtp | Where-Object { $_ -ieq $subject }).Count -eq 0) {
                    Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'Failed' `
                                 -PrimaryBefore $currentPrimary -PrimaryAfter $afterPrimary `
                                 -Detail 'Exchange accepted the change but the alias is not on the mailbox.'

                    Write-Warning "$mailboxName : alias $subject is not present after the write."
                }
                else {
                    Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'AliasAdded' `
                                 -PrimaryBefore $currentPrimary -PrimaryAfter $afterPrimary `
                                 -Detail "Primary left as $afterPrimary"
                }
            }
            catch {
                Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'Failed' `
                             -PrimaryBefore $currentPrimary -Detail $_.Exception.Message

                Write-Warning "$mailboxName : $($_.Exception.Message)"
            }
        }
        else {
            Write-Result -UserPrincipalName $upn -Mailbox $mailboxName -Address $subject -Status 'WhatIf' `
                         -PrimaryBefore $currentPrimary -PrimaryAfter $currentPrimary `
                         -Detail "Would add $subject as an alias, leaving $currentPrimary as the primary"
        }

        #endregion Add the old address as an alias ----------------------------
    }
}

Write-Progress -Activity "Updating mailbox addresses ($Action)" -Completed

#endregion Process ------------------------------------------------------------

#region Report ----------------------------------------------------------------

# -WhatIf:$false so a rehearsal still produces its report - the log is the whole
# point of a dry run.
$log | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8 -WhatIf:$false

Write-Host ''
Write-Host '--- Summary ---' -ForegroundColor Green

$log | Group-Object Status | Sort-Object Name | ForEach-Object {
    Write-Host ("  {0,-20} {1}" -f $_.Name, $_.Count)
}

Write-Host ''
Write-Host "  Log : $((Resolve-Path -Path $LogPath).Path)" -ForegroundColor Green

$failures = @($log | Where-Object { $_.Status -eq 'Failed' })

if ($failures.Count -gt 0) {
    Write-Host "  $($failures.Count) mailbox(es) failed - see the Detail column." -ForegroundColor Red
}

$needsAttention = @($log | Where-Object { $_.Status -in @('NotFound', 'Mismatch', 'Conflict', 'DomainNotAccepted', 'PolicyManaged', 'DirSynced', 'Refused') })

if ($needsAttention.Count -gt 0) {
    Write-Host "  $($needsAttention.Count) row(s) were left alone and need a look - see the Status column." -ForegroundColor Yellow
}

Disconnect-ExchangeOnline -Confirm:$false | Out-Null

Write-Host ''
Write-Host 'Done. Re-running is safe: anything already correct is reported and skipped.' -ForegroundColor Green

#endregion Report -------------------------------------------------------------
