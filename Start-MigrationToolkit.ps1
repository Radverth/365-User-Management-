<#
.SYNOPSIS
    Step-by-step guided menu for the migration scripts.

.DESCRIPTION
    Start here if you would rather answer questions than remember parameters. Every
    task in this toolkit is reachable from the menu, and each one explains what it
    is about to do before it does it.

    Anything that changes your tenant is rehearsed first: the wizard runs it in
    preview mode, shows you the results, and only makes real changes after you
    confirm. Reports never change anything.

    The equivalent command line is printed for every task, so you can repeat a run
    later without the menu.

    Runs on Windows, Linux and macOS. A windowed interface would not - PowerShell's
    WinForms and WPF are Windows-only - so this uses the terminal instead.

.EXAMPLE
    .\Start-MigrationToolkit.ps1
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$script:Root = $PSScriptRoot

#region Console helpers -------------------------------------------------------

function Write-Banner {
    param([Parameter(Mandatory)][string]$Text)

    $line = '=' * 62

    Write-Host ''
    Write-Host $line -ForegroundColor Cyan
    Write-Host ("  " + $Text) -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan
    Write-Host ''
}

function Write-Step {
    param([Parameter(Mandatory)][string]$Text)

    Write-Host ''
    Write-Host "-- $Text " -ForegroundColor White -NoNewline
    Write-Host ('-' * [Math]::Max(0, 58 - $Text.Length)) -ForegroundColor DarkGray
    Write-Host ''
}

function Write-Explain {
    # AllowEmptyString so a blank entry can be used as a paragraph break; a
    # mandatory [string[]] rejects empty elements otherwise.
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string[]]$Lines
    )

    foreach ($line in $Lines) { Write-Host "   $line" -ForegroundColor DarkGray }
    Write-Host ''
}

function Read-Line {
    <#  Read-Host returns $null once input runs out - piped input exhausted, or the
        console closed. $null does not match an empty-string test and blows up on
        .Trim(), so it is converted into something the top level can catch and exit
        on cleanly. #>
    param([Parameter(Mandatory)][string]$Prompt)

    $value = Read-Host -Prompt $Prompt

    if ($null -eq $value) { throw [System.IO.EndOfStreamException]::new('Input ended.') }

    return $value
}

function Read-Secret {
    <#  Password entry. Read-Host -AsSecureString needs a real console: with input
        redirected it reads nothing and silently ends the script, so that case falls
        back to a visible prompt and says so rather than dying quietly. #>
    param([Parameter(Mandatory)][string]$Prompt)

    if ([Console]::IsInputRedirected) {

        Write-Host '   Input is redirected, so the password cannot be hidden as you type.' -ForegroundColor Yellow

        return (ConvertTo-SecureString -String (Read-Line -Prompt $Prompt) -AsPlainText -Force)
    }

    $value = Read-Host -Prompt $Prompt -AsSecureString

    if ($null -eq $value) { throw [System.IO.EndOfStreamException]::new('Input ended.') }

    return $value
}

function Read-Text {
    <#  Read-Host with a retry limit, so a closed or exhausted input stream ends the
        wizard instead of spinning forever. #>
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [string]$Default,
        [switch]$AllowEmpty,
        [scriptblock]$Validate,
        [string]$ValidationMessage = 'That does not look right - please try again.'
    )

    $label = if ($Default) { "$Prompt [$Default]" } else { $Prompt }

    for ($attempt = 0; $attempt -lt 10; $attempt++) {

        $value = Read-Line -Prompt $label

        if ([string]::IsNullOrWhiteSpace($value) -and $Default) { $value = $Default }

        if ([string]::IsNullOrWhiteSpace($value)) {

            if ($AllowEmpty) { return '' }

            Write-Host '   This one is required.' -ForegroundColor Yellow
            continue
        }

        $value = $value.Trim()

        if ($Validate -and -not (& $Validate $value)) {
            Write-Host "   $ValidationMessage" -ForegroundColor Yellow
            continue
        }

        return $value
    }

    throw 'No usable input received. Exiting.'
}

function Read-Choice {
    <#  Numbered menu. Returns the Key of the chosen option. #>
    param(
        [Parameter(Mandatory)][array]$Options,
        [string]$Prompt = 'Choose'
    )

    for ($attempt = 0; $attempt -lt 10; $attempt++) {

        $index = 1

        foreach ($option in $Options) {
            Write-Host ("   {0}. {1}" -f $index, $option.Label) -ForegroundColor White

            if ($option.Detail) { Write-Host ("      {0}" -f $option.Detail) -ForegroundColor DarkGray }

            $index++
        }

        Write-Host ''

        $answer = Read-Line -Prompt "$Prompt [1-$($Options.Count)]"

        if ($answer -match '^\s*$') {
            Write-Host '   Please pick a number.' -ForegroundColor Yellow
            Write-Host ''
            continue
        }

        $answer = $answer.Trim()

        if ($answer -match '^\d+$') {
            $picked = [int]$answer

            if ($picked -ge 1 -and $picked -le $Options.Count) { return $Options[$picked - 1].Key }
        }

        Write-Host "   Enter a number between 1 and $($Options.Count)." -ForegroundColor Yellow
        Write-Host ''
    }

    throw 'No usable input received. Exiting.'
}

function Read-YesNo {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [bool]$Default = $false
    )

    $hint = if ($Default) { 'Y/n' } else { 'y/N' }

    for ($attempt = 0; $attempt -lt 10; $attempt++) {

        $answer = Read-Line -Prompt "$Prompt [$hint]"

        if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }

        switch -Regex ($answer.Trim()) {
            '^(y|yes)$' { return $true }
            '^(n|no)$'  { return $false }
            default     { Write-Host '   Please answer y or n.' -ForegroundColor Yellow }
        }
    }

    throw 'No usable input received. Exiting.'
}

function Show-Command {
    <#  Prints the equivalent command line, so the run can be repeated without the
        wizard and so there is no mystery about what is being executed. #>
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][hashtable]$Parameters
    )

    $name = Split-Path -Path $ScriptPath -Leaf

    Write-Host '   Equivalent command:' -ForegroundColor DarkGray
    Write-Host "     .\$name" -ForegroundColor DarkCyan -NoNewline

    foreach ($key in ($Parameters.Keys | Sort-Object)) {

        $value = $Parameters[$key]

        if ($value -is [switch] -or $value -is [bool]) {
            if ($value) { Write-Host " -$key" -ForegroundColor DarkCyan -NoNewline }
        }
        elseif ($value -is [System.Security.SecureString]) {
            Write-Host " -$key <hidden>" -ForegroundColor DarkCyan -NoNewline
        }
        elseif ($value -is [array]) {
            Write-Host " -$key $($value -join ',')" -ForegroundColor DarkCyan -NoNewline
        }
        elseif ("$value" -match '\s') {
            Write-Host " -$key `"$value`"" -ForegroundColor DarkCyan -NoNewline
        }
        else {
            Write-Host " -$key $value" -ForegroundColor DarkCyan -NoNewline
        }
    }

    Write-Host ''
    Write-Host ''
}

function Invoke-ToolkitScript {
    <#  Runs one of the scripts, reporting failure in plain language rather than a
        stack trace. Returns $true on success. #>
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][hashtable]$Parameters
    )

    if (-not (Test-Path -Path $ScriptPath)) {
        Write-Host "   Could not find $ScriptPath" -ForegroundColor Red
        Write-Host '   Run this from inside a full copy of the repository.' -ForegroundColor Red
        return $false
    }

    try {
        & $ScriptPath @Parameters
        return $true
    }
    catch {
        Write-Host ''
        Write-Host '   That did not work:' -ForegroundColor Red

        foreach ($line in ($_.Exception.Message -split "`n")) {
            Write-Host "     $($line.TrimEnd())" -ForegroundColor Red
        }

        Write-Host ''
        return $false
    }
}

function Invoke-WithRehearsal {
    <#  Preview, show the result, then ask before doing it for real. Used for
        everything that changes the tenant. #>
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][hashtable]$Parameters,
        [Parameter(Mandatory)][string]$Description
    )

    Write-Step 'Step 1 of 2: rehearsal'
    Write-Explain @(
        'Nothing is changed yet. This shows exactly what would happen so you',
        'can check it before committing.'
    )

    Show-Command -ScriptPath $ScriptPath -Parameters $Parameters

    $preview = $Parameters.Clone()
    $preview['WhatIf'] = $true

    if (-not (Invoke-ToolkitScript -ScriptPath $ScriptPath -Parameters $preview)) { return }

    Write-Step 'Step 2 of 2: apply the changes'
    Write-Explain @(
        "About to: $Description",
        'Check the preview above, and the log file it produced, before saying yes.'
    )

    if (-not (Read-YesNo -Prompt '   Go ahead and make these changes?' -Default $false)) {
        Write-Host ''
        Write-Host '   Nothing was changed.' -ForegroundColor Yellow
        return
    }

    Write-Host ''
    [void](Invoke-ToolkitScript -ScriptPath $ScriptPath -Parameters $Parameters)
}

#endregion Console helpers ----------------------------------------------------

#region Shared prompts --------------------------------------------------------

function Test-LooksLikeUrl { param($Value) return ($Value -match '^https://[^\s]+$') }
function Test-LooksLikeTenant { param($Value) return ($Value -match '^[A-Za-z0-9-]+\.[A-Za-z0-9.]+$') }
function Test-LooksLikeGuid { param($Value) return ($Value -match '^[0-9a-fA-F-]{36}$') }

function Get-SharePointAuth {
    <#  Asks how to sign in once, and returns the parameters to splat into the
        SharePoint scripts. #>

    Write-Step 'How should this sign in?'

    $choice = Read-Choice -Options @(
        @{ Key = 'interactive'; Label = 'Sign in as me'
           Detail = 'Simplest. A browser opens. You will only see sites you administer.' }
        @{ Key = 'appOnly'; Label = 'Use the app certificate'
           Detail = 'Reaches every site in the tenant. Needs the certificate set up first (option 5 on the main menu).' }
    )

    $auth = @{}

    if ($choice -eq 'interactive') {

        Write-Explain @(
            'PnP needs an app registration even to sign in as you. If you do not',
            'have a client ID, see the README section "Entra app registration".'
        )

        $auth['ClientId'] = Read-Text -Prompt '   Client ID' `
            -Validate ${function:Test-LooksLikeGuid} `
            -ValidationMessage 'A client ID looks like 66792cc7-5f4e-4657-9448-1bd119175f55.'

        return $auth
    }

    $auth['ClientId'] = Read-Text -Prompt '   Client ID' `
        -Validate ${function:Test-LooksLikeGuid} `
        -ValidationMessage 'A client ID looks like 66792cc7-5f4e-4657-9448-1bd119175f55.'

    $auth['Tenant'] = Read-Text -Prompt '   Tenant name (e.g. contoso.onmicrosoft.com)' `
        -Validate ${function:Test-LooksLikeTenant} `
        -ValidationMessage 'That should look like contoso.onmicrosoft.com.'

    Write-Host ''
    Write-Explain @('The .pfx file created when you set up the certificate.')

    $auth['CertificatePath'] = Read-Text -Prompt '   Path to the .pfx file' `
        -Validate { param($v) Test-Path -Path $v } `
        -ValidationMessage 'No file at that path. Check the location and try again.'

    if (Read-YesNo -Prompt '   Is the .pfx password protected?' -Default $true) {
        $auth['CertificatePassword'] = Read-Secret -Prompt '   Certificate password'
    }

    return $auth
}

function Get-SiteSelection {
    <#  Returns the parameters describing which sites to act on. #>
    param([switch]$AllowAllSites)

    Write-Step 'Which sites?'

    $options = @()

    if ($AllowAllSites) {
        $options += @{ Key = 'all'; Label = 'Every site in the tenant'
                       Detail = 'Takes a while on a large tenant - each site is opened in turn.' }
    }

    $options += @{ Key = 'one'; Label = 'One site I will type in' }
    $options += @{ Key = 'csv'; Label = 'A list of sites from a spreadsheet'
                   Detail = 'A .csv file with a column headed SiteUrl.' }

    $choice = Read-Choice -Options $options

    switch ($choice) {

        'all' {
            Write-Host ''
            Write-Explain @(
                'Your tenant admin address is your SharePoint address with -admin',
                'added: https://contoso.sharepoint.com becomes',
                'https://contoso-admin.sharepoint.com'
            )

            return @{
                AllSites       = $true
                TenantAdminUrl = Read-Text -Prompt '   Tenant admin URL' `
                    -Validate ${function:Test-LooksLikeUrl} `
                    -ValidationMessage 'That should start with https://'
            }
        }

        'one' {
            return @{
                SiteUrl = @(Read-Text -Prompt '   Site URL' `
                    -Validate ${function:Test-LooksLikeUrl} `
                    -ValidationMessage 'That should start with https://')
            }
        }

        'csv' {
            return @{
                SitesCsvPath = Read-Text -Prompt '   Path to the .csv file' `
                    -Validate { param($v) Test-Path -Path $v } `
                    -ValidationMessage 'No file at that path. Check the location and try again.'
            }
        }
    }
}

function Get-OutputPath {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [Parameter(Mandatory)][string]$Default
    )

    $path = Read-Text -Prompt "   $Prompt" -Default $Default

    if ((Test-Path -Path $path) -and -not (Read-YesNo -Prompt "   $path exists. Overwrite it?" -Default $true)) {
        return Get-OutputPath -Prompt $Prompt -Default $Default
    }

    return $path
}

#endregion Shared prompts -----------------------------------------------------

#region Tasks -----------------------------------------------------------------

# The submodules the scripts' #Requires lines actually name. Installing these is
# far quicker than the Microsoft.Graph meta-module, which pulls in around forty.
$script:RequiredModules = @(
    @{ Name = 'Microsoft.Graph.Users';                        For = 'guest scripts' }
    @{ Name = 'Microsoft.Graph.Groups';                       For = 'guest scripts' }
    @{ Name = 'Microsoft.Graph.Identity.SignIns';             For = 'guest import (invitations)' }
    @{ Name = 'Microsoft.Graph.Identity.DirectoryManagement'; For = 'guest export' }
    @{ Name = 'PnP.PowerShell';                               For = 'SharePoint scripts' }
)

function Get-ModuleStatus {
    <#  One row per required module, with the installed version when present. #>

    foreach ($required in $script:RequiredModules) {

        $found = Get-Module -ListAvailable -Name $required.Name |
                    Sort-Object Version -Descending |
                    Select-Object -First 1

        [PSCustomObject]@{
            Name      = $required.Name
            For       = $required.For
            Installed = [bool]$found
            Version   = if ($found) { "$($found.Version)" } else { '' }
        }
    }
}

function Install-MissingModules {
    <#  Installs into the current user's module path, so no administrator rights are
        needed. Returns the names that are still missing afterwards. #>
    param([Parameter(Mandatory)][string[]]$Names)

    Write-Host ''
    Write-Explain @(
        'These come from the PowerShell Gallery, Microsoft''s public module',
        'repository. They install for your user account only, so no administrator',
        'rights are needed.',
        '',
        'This can take a few minutes.'
    )

    $repository = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue

    if ($repository -and $repository.InstallationPolicy -ne 'Trusted') {
        Write-Host '   The PowerShell Gallery is marked untrusted on this machine, so the' -ForegroundColor DarkGray
        Write-Host '   install is run with -Force to avoid a prompt for each module.' -ForegroundColor DarkGray
        Write-Host ''
    }

    foreach ($name in $Names) {

        Write-Host "   Installing $name ..." -ForegroundColor Cyan

        try {
            Install-Module -Name $name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop

            Write-Host "     done" -ForegroundColor Green
        }
        catch {
            Write-Host "     failed: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "     Try it by hand: Install-Module $name -Scope CurrentUser" -ForegroundColor Yellow
        }
    }

    Write-Host ''

    return @(Get-ModuleStatus | Where-Object { -not $_.Installed -and $_.Name -in $Names } |
                Select-Object -ExpandProperty Name)
}

function Invoke-CheckSetup {

    Write-Banner 'Check what is installed'

    $modules = @(Get-ModuleStatus)

    $rows = foreach ($module in $modules) {
        [PSCustomObject]@{
            Component = $module.Name
            Status    = if ($module.Installed) { "v$($module.Version)" } else { 'MISSING' }
            Needed_by = $module.For
        }
    }

    $rows | Format-Table -AutoSize | Out-String -Width 100 | Write-Host

    $psVersion = $PSVersionTable.PSVersion

    if ($psVersion.Major -lt 7) {
        Write-Host "   PowerShell is v$psVersion. Version 7 or later is recommended." -ForegroundColor Yellow
    }
    else {
        Write-Host "   PowerShell v$psVersion - fine." -ForegroundColor Green
    }

    # The scripts dot-source shared helpers, so a partial copy fails confusingly.
    $missingFiles = @()

    foreach ($relative in @('Guests/Export-GuestPermissions.ps1', 'Guests/Import-GuestPermissions.ps1',
                            'SharePoint/Get-SiteOwners.ps1', 'SharePoint/Set-SiteMembersToViewers.ps1',
                            'Setup/New-AppOnlyCertificate.ps1', 'Common/InputCsv.ps1', 'Common/PnPConnect.ps1')) {

        if (-not (Test-Path -Path (Join-Path -Path $script:Root -ChildPath $relative))) { $missingFiles += $relative }
    }

    if ($missingFiles.Count -eq 0) {
        Write-Host '   All toolkit files present.' -ForegroundColor Green
    }
    else {
        Write-Host '   Some toolkit files are missing:' -ForegroundColor Red

        foreach ($file in $missingFiles) { Write-Host "     $file" -ForegroundColor Red }

        Write-Host '   Ask for a complete copy of the toolkit folder - the scripts share files in Common/.' -ForegroundColor Yellow
    }

    $missingModules = @($modules | Where-Object { -not $_.Installed } | Select-Object -ExpandProperty Name)

    if ($missingModules.Count -eq 0) {

        Write-Host '   All required modules are installed.' -ForegroundColor Green
        Write-Host ''

        # An older PnP has no -PersistLogin, which means a sign-in prompt at every
        # single site. Worth offering, but not worth churning if it is already fine.
        if (Read-YesNo -Prompt '   Update PnP.PowerShell to the latest version? (older versions ask you to sign in at every site)' -Default $false) {

            Write-Host '   Updating PnP.PowerShell ...' -ForegroundColor Cyan

            try {
                Update-Module -Name PnP.PowerShell -Force -ErrorAction Stop
                Write-Host '     done' -ForegroundColor Green
            }
            catch {
                Write-Host "     failed: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host '     Try it by hand: Update-Module PnP.PowerShell' -ForegroundColor Yellow
            }
        }

        Write-Host ''
        return
    }

    Write-Host ''
    Write-Host "   $($missingModules.Count) module(s) need installing:" -ForegroundColor Yellow

    foreach ($name in $missingModules) { Write-Host "     $name" -ForegroundColor Yellow }

    Write-Host ''

    if (-not (Read-YesNo -Prompt '   Install them now?' -Default $true)) {

        Write-Host ''
        Write-Host '   Nothing installed. To do it yourself:' -ForegroundColor Yellow

        foreach ($name in $missingModules) {
            Write-Host "     Install-Module $name -Scope CurrentUser" -ForegroundColor Yellow
        }

        Write-Host ''
        return
    }

    $stillMissing = Install-MissingModules -Names $missingModules

    if ($stillMissing.Count -eq 0) {
        Write-Host '   All required modules are now installed.' -ForegroundColor Green
    }
    else {
        Write-Host '   Still missing after the attempt:' -ForegroundColor Red

        foreach ($name in $stillMissing) { Write-Host "     $name" -ForegroundColor Red }
    }

    Write-Host ''
}

function Invoke-ExportGuests {

    Write-Banner 'Export guest users from the old tenant'

    Write-Explain @(
        'Produces a spreadsheet of every guest and the groups they belong to.',
        'Nothing is changed. Run this against the tenant you are moving away from.'
    )

    Write-Step 'Where should the spreadsheet go?'

    $parameters = @{
        OutputPath = Get-OutputPath -Prompt 'File to create' -Default './TenantA_GuestPermissions.csv'
    }

    Write-Step 'Which guests?'

    Write-Explain @(
        'Guests who are blocked from signing in, or who never accepted their',
        'original invitation, are left out unless you ask for them.'
    )

    if (Read-YesNo -Prompt '   Include guests who are blocked from signing in?' -Default $false) {
        $parameters['IncludeDisabledGuests'] = $true
    }

    if (Read-YesNo -Prompt '   Include guests who never accepted their invitation?' -Default $false) {
        $parameters['IncludePendingAcceptance'] = $true
    }

    Write-Step 'Running the export'
    Write-Explain @('A browser will open for you to sign in to the OLD tenant.')

    Show-Command -ScriptPath (Join-Path $script:Root 'Guests/Export-GuestPermissions.ps1') -Parameters $parameters

    [void](Invoke-ToolkitScript -ScriptPath (Join-Path $script:Root 'Guests/Export-GuestPermissions.ps1') -Parameters $parameters)
}

function Invoke-ImportGuests {

    Write-Banner 'Create those guests in the new tenant'

    Write-Explain @(
        'Reads the spreadsheet from the export and recreates the guests and their',
        'group memberships in the new tenant.',
        '',
        'It never removes anything, and running it twice is safe - anyone already',
        'there is left alone.',
        '',
        'The groups must already exist in the new tenant with the same names.'
    )

    Write-Step 'Which spreadsheet?'

    $parameters = @{
        InputPath = Read-Text -Prompt '   Path to the exported .csv' -Default './TenantA_GuestPermissions.csv' `
            -Validate { param($v) Test-Path -Path $v } `
            -ValidationMessage 'No file at that path. Run the export first, or check the location.'
    }

    Write-Step 'Should the guests be emailed?'

    Write-Explain @(
        'By default guests are created quietly and receive nothing, which is',
        'usually what you want until the migration is finished.'
    )

    if (Read-YesNo -Prompt '   Send each new guest an invitation email?' -Default $false) {

        $parameters['SendInvitationMessage'] = $true

        $message = Read-Text -Prompt '   Message to include (press Enter to skip)' -AllowEmpty

        if ($message) { $parameters['CustomInvitationMessage'] = $message }
    }

    Write-Step 'How many?'

    Write-Explain @(
        'Doing a handful first is a good way to check the result before',
        'committing to everyone.'
    )

    if (Read-YesNo -Prompt '   Start with just a few guests as a trial?' -Default $true) {

        $parameters['MaxGuests'] = [int](Read-Text -Prompt '   How many?' -Default '5' `
            -Validate { param($v) $v -match '^\d+$' -and [int]$v -gt 0 } `
            -ValidationMessage 'Enter a whole number greater than zero.')
    }

    $parameters['LogPath'] = Get-OutputPath -Prompt 'Where should the results log go?' -Default './TenantB_GuestImport_Log.csv'

    Write-Explain @('A browser will open for you to sign in to the NEW tenant.')

    $count = if ($parameters.ContainsKey('MaxGuests')) { "up to $($parameters['MaxGuests']) guest(s)" } else { 'every guest in the file' }
    $email = if ($parameters.ContainsKey('SendInvitationMessage')) { 'and email them an invitation' } else { 'without emailing them' }

    Invoke-WithRehearsal -ScriptPath (Join-Path $script:Root 'Guests/Import-GuestPermissions.ps1') `
                         -Parameters $parameters `
                         -Description "create $count in the new tenant $email, and add them to their groups"
}

function Invoke-SiteOwners {

    Write-Banner 'Report who owns your SharePoint sites'

    Write-Explain @(
        'Produces a spreadsheet of site owners. Nothing is changed.',
        '',
        'Owner means three different things in SharePoint, and all three are',
        'listed, so one person can appear more than once for the same site.'
    )

    $parameters = Get-SiteSelection -AllowAllSites
    $parameters += Get-SharePointAuth

    Write-Step 'Where should the spreadsheet go?'

    $parameters['OutputPath'] = Get-OutputPath -Prompt 'File to create' -Default './SharePoint_SiteOwners.csv'

    Write-Step 'Running the report'

    Show-Command -ScriptPath (Join-Path $script:Root 'SharePoint/Get-SiteOwners.ps1') -Parameters $parameters

    [void](Invoke-ToolkitScript -ScriptPath (Join-Path $script:Root 'SharePoint/Get-SiteOwners.ps1') -Parameters $parameters)
}

function Invoke-MembersToViewers {

    Write-Banner 'Change site members to view-only'

    Write-Explain @(
        'Moves people out of a site''s Members group and into its Visitors group,',
        'so they can read the site but no longer edit it.',
        '',
        'People are added to Visitors before being removed from Members, so nobody',
        'is ever left without access.'
    )

    $parameters = Get-SiteSelection
    $parameters += Get-SharePointAuth

    Write-Step 'Who should be moved?'

    $who = Read-Choice -Options @(
        @{ Key = 'guests'; Label = 'Guests only'
           Detail = 'External people. Your own staff keep their current access.' }
        @{ Key = 'everyone'; Label = 'Everyone listed as a member'
           Detail = 'Guests and your own staff.' }
    )

    if ($who -eq 'everyone') { $parameters['IncludeInternalUsers'] = $true }

    Write-Step 'Teams sites'

    Write-Explain @(
        'On a site connected to a Team, the Members list contains the Team itself',
        '(shown as "<Site> Members") as well as any individuals.',
        '',
        'Moving that entry makes the whole Team read-only on this site in one go,',
        'including people who join later. It does not affect Teams chat, the',
        'shared mailbox or the calendar.'
    )

    if (Read-YesNo -Prompt '   Move the Team entry as well?' -Default $false) {
        $parameters['IncludeSecurityGroups'] = $true
    }

    Write-Step 'Anyone to leave alone?'

    Write-Explain @('For example an administrator or service account that must keep editing.')

    $exclude = Read-Text -Prompt '   Email addresses to skip, separated by commas (press Enter for none)' -AllowEmpty

    if ($exclude) {
        $parameters['ExcludeLogin'] = @($exclude -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    $parameters['OutputPath'] = Get-OutputPath -Prompt 'Where should the results log go?' -Default './SharePoint_MembersToViewers_Log.csv'

    $scope = if ($parameters.ContainsKey('IncludeInternalUsers')) { 'everyone in the Members group' } else { 'guests in the Members group' }
    $team  = if ($parameters.ContainsKey('IncludeSecurityGroups')) { ', including the connected Team' } else { '' }

    Invoke-WithRehearsal -ScriptPath (Join-Path $script:Root 'SharePoint/Set-SiteMembersToViewers.ps1') `
                         -Parameters $parameters `
                         -Description "move $scope$team to view-only"
}

function Invoke-CreateCertificate {

    Write-Banner 'Set up a certificate for full site access'

    Write-Explain @(
        'Signing in as yourself only reaches sites you administer. A certificate',
        'lets the scripts sign in as the application instead, which reaches every',
        'site in the tenant.',
        '',
        'This creates the certificate. You then upload one of the two files it',
        'produces to your app registration - it will tell you how.'
    )

    $parameters = @{
        CommonName = Read-Text -Prompt '   A name for the certificate' -Default 'PnP Migration'
        OutPath    = Read-Text -Prompt '   Folder to save it in' -Default './certs'
        ValidYears = [int](Read-Text -Prompt '   How many years should it last?' -Default '1' `
            -Validate { param($v) $v -match '^\d+$' -and [int]$v -ge 1 -and [int]$v -le 30 } `
            -ValidationMessage 'Enter a whole number between 1 and 30.')
    }

    Write-Step 'Password'

    Write-Explain @(
        'The .pfx file holds a private key. Anyone who has it can sign in as the',
        'application, so protecting it with a password is strongly recommended.'
    )

    if (Read-YesNo -Prompt '   Protect it with a password?' -Default $true) {
        $parameters['CertificatePassword'] = Read-Secret -Prompt '   Choose a password'
    }

    if (Read-YesNo -Prompt '   Also install it on this machine, so you can use the thumbprint?' -Default $true) {
        $parameters['Install'] = $true
    }

    Write-Step 'Creating the certificate'

    [void](Invoke-ToolkitScript -ScriptPath (Join-Path $script:Root 'Setup/New-AppOnlyCertificate.ps1') -Parameters $parameters)
}

#endregion Tasks --------------------------------------------------------------

Write-Banner 'Microsoft 365 Migration Toolkit'

Write-Host '   Answer the questions and this will run the right script for you.' -ForegroundColor DarkGray
Write-Host '   Anything that makes changes is previewed first, and asks before' -ForegroundColor DarkGray
Write-Host '   committing. Reports never change anything.' -ForegroundColor DarkGray

try {

for ($loop = 0; $loop -lt 100; $loop++) {

    Write-Step 'What would you like to do?'

    $task = Read-Choice -Options @(
        @{ Key = 'export';  Label = 'Export guest users from the old tenant';   Detail = 'Creates a spreadsheet. Changes nothing.' }
        @{ Key = 'import';  Label = 'Create those guests in the new tenant';    Detail = 'Makes changes. Previewed first.' }
        @{ Key = 'owners';  Label = 'Report who owns your SharePoint sites';    Detail = 'Creates a spreadsheet. Changes nothing.' }
        @{ Key = 'viewers'; Label = 'Change site members to view-only';         Detail = 'Makes changes. Previewed first.' }
        @{ Key = 'cert';    Label = 'Set up a certificate for full site access'; Detail = 'Do this once, if you need every site.' }
        @{ Key = 'check';   Label = 'Check and install what is needed'; Detail = 'Offers to install any missing PowerShell modules.' }
        @{ Key = 'quit';    Label = 'Quit' }
    )

    switch ($task) {
        'export'  { Invoke-ExportGuests }
        'import'  { Invoke-ImportGuests }
        'owners'  { Invoke-SiteOwners }
        'viewers' { Invoke-MembersToViewers }
        'cert'    { Invoke-CreateCertificate }
        'check'   { Invoke-CheckSetup }
        'quit'    {
            Write-Host ''
            Write-Host '   Done.' -ForegroundColor Green
            Write-Host ''
            return
        }
    }
}

}
catch [System.IO.EndOfStreamException] {
    Write-Host ''
    Write-Host '   Input ended - exiting.' -ForegroundColor Yellow
    Write-Host ''
    return
}
