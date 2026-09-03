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

# Shared CSV input handling, so the wizard can validate a site list at the moment
# it is chosen rather than leaving it to the script it launches.
$inputCsvPath = Join-Path -Path $script:Root -ChildPath 'Common/InputCsv.ps1'

if (Test-Path -Path $inputCsvPath) { . $inputCsvPath }

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
    <#  Numbered menu. Returns the Key of the chosen option.

        An entry carrying a Header instead of a Label prints as a section heading
        and is not numbered, so a long menu can be grouped without the numbering
        skipping. An entry may also carry a Tag, printed on the right, which is
        where read-only versus makes-changes is shown. #>
    param(
        [Parameter(Mandatory)][array]$Options,
        [string]$Prompt = 'Choose'
    )

    $selectable = @($Options | Where-Object { -not $_.Header })

    for ($attempt = 0; $attempt -lt 10; $attempt++) {

        $index = 1
        $printedAny = $false

        foreach ($option in $Options) {

            if ($option.Header) {
                # Separator between groups, but not above the first one - the
                # caller has already spaced the menu away from what precedes it.
                if ($printedAny) { Write-Host '' }

                Write-Host ("   {0}" -f $option.Header.ToUpper()) -ForegroundColor Cyan

                if ($option.Note) { Write-Host ("   {0}" -f $option.Note) -ForegroundColor DarkGray }

                $printedAny = $true
                continue
            }

            $line = "    {0,2}  {1}" -f $index, $option.Label

            Write-Host $line -ForegroundColor White -NoNewline

            if ($option.Tag) {
                $pad = [Math]::Max(1, 56 - $line.Length)
                $colour = if ($option.Tag -match 'change|email') { 'Yellow' } else { 'DarkGray' }

                Write-Host ((' ' * $pad) + $option.Tag) -ForegroundColor $colour
            }
            else {
                Write-Host ''
            }

            if ($option.Detail) { Write-Host ("        {0}" -f $option.Detail) -ForegroundColor DarkGray }

            $printedAny = $true
            $index++
        }

        Write-Host ''

        $answer = Read-Line -Prompt "$Prompt [1-$($selectable.Count)]"

        if ($answer -match '^\s*$') {
            Write-Host '   Please pick a number.' -ForegroundColor Yellow
            Write-Host ''
            continue
        }

        $answer = $answer.Trim()

        if ($answer -match '^\d+$') {
            $picked = [int]$answer

            if ($picked -ge 1 -and $picked -le $selectable.Count) { return $selectable[$picked - 1].Key }
        }

        Write-Host "   Enter a number between 1 and $($selectable.Count)." -ForegroundColor Yellow
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
        [string]$ScriptPath,

        # For a cmdlet rather than one of the toolkit's scripts.
        [string]$CommandName,

        [Parameter(Mandatory)][hashtable]$Parameters
    )

    $name = if ($CommandName) { $CommandName } else { '.\' + (Split-Path -Path $ScriptPath -Leaf) }

    Write-Host '   Equivalent command:' -ForegroundColor DarkGray
    Write-Host "     $name" -ForegroundColor DarkCyan -NoNewline

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
           Detail = 'Reaches every site in the tenant. Needs the certificate created first - "Create a certificate for full site access" on the main menu.' }
    )

    $auth = @{}

    if ($choice -eq 'interactive') {

        Write-Explain @(
            'PnP needs an app registration even to sign in as you. If you do not',
            'have a client ID yet, quit and run "Register the app in Entra ID"',
            'from the main menu first - it creates one and prints the ID.'
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
                       Detail = 'Each site is opened in turn, so allow time on a large tenant.' }
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

            Write-Host ''
            Write-Explain @(
                'The file needs one column headed SiteUrl, and one site per row:',
                '',
                '    SiteUrl',
                '    https://contoso.sharepoint.com/sites/Marketing',
                '    https://contoso.sharepoint.com/sites/Projects',
                '',
                'Other columns are ignored, so a site list exported from elsewhere',
                'usually works as-is.'
            )

            for ($attempt = 0; $attempt -lt 5; $attempt++) {

                $path = Read-Text -Prompt '   Path to the .csv file' `
                    -Validate { param($v) Test-Path -Path $v } `
                    -ValidationMessage 'No file at that path. Check the location and try again.'

                # Read it now rather than letting the script fail later: a beginner
                # who picked the wrong file should find out here, and seeing the
                # count confirms it is the list they meant.
                try {
                    $rows = Import-InputCsv -Path $path -Expected 'a site list with a SiteUrl column' -RequiredColumns @('SiteUrl')

                    $found = @($rows | Select-Object -ExpandProperty SiteUrl |
                                Where-Object { $_ } | ForEach-Object { $_.Trim() })

                    if ($found.Count -eq 0) {
                        Write-Host '   That file has a SiteUrl column, but every row is empty.' -ForegroundColor Yellow
                        continue
                    }

                    Write-Host ''
                    Write-Host "   Found $($found.Count) site(s) in that file:" -ForegroundColor Green

                    foreach ($site in ($found | Select-Object -First 5)) {
                        Write-Host "     $site" -ForegroundColor DarkGray
                    }

                    if ($found.Count -gt 5) {
                        Write-Host "     ... and $($found.Count - 5) more" -ForegroundColor DarkGray
                    }

                    Write-Host ''

                    if (Read-YesNo -Prompt '   Is that the right list?' -Default $true) {
                        return @{ SitesCsvPath = $path }
                    }
                }
                catch {
                    Write-Host ''
                    Write-Host '   That file could not be read as a site list:' -ForegroundColor Yellow

                    foreach ($line in ($_.Exception.Message -split "`n")) {
                        Write-Host "     $($line.TrimEnd())" -ForegroundColor Yellow
                    }

                    Write-Host ''
                }
            }

            throw 'No usable site list was given. Exiting.'
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
                            'Setup/New-AppOnlyCertificate.ps1', 'Reports/Get-M365UserPermissionsReport.ps1',
                            'Common/InputCsv.ps1', 'Common/PnPConnect.ps1')) {

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

function Invoke-UserPermissionsReport {

    Write-Banner 'List all users and their group access'

    Write-Explain @(
        'Produces a spreadsheet of every user in the tenant - staff and guests',
        'alike - with the groups each one belongs to. Nothing is changed.',
        '',
        'This shows EFFECTIVE access, so a group someone is in only through',
        'another group is included. That is what you want for a permissions',
        'review, and is why it differs from the guest export, which records',
        'direct memberships only.'
    )

    Write-Step 'Where should the spreadsheet go?'

    $parameters = @{
        OutputPath = Get-OutputPath -Prompt 'File to create' -Default './M365_User_Permissions_Report.csv'
    }

    Write-Step 'Which tenant?'

    Write-Explain @(
        'Leave this blank unless your account exists in more than one tenant',
        'and you need to pick a particular one.'
    )

    $tenant = Read-Text -Prompt '   Tenant name (press Enter to be asked at sign-in)' -AllowEmpty

    if ($tenant) { $parameters['TenantId'] = $tenant }

    Write-Step 'Running the report'
    Write-Explain @('A browser will open for you to sign in.')

    Show-Command -ScriptPath (Join-Path $script:Root 'Reports/Get-M365UserPermissionsReport.ps1') -Parameters $parameters

    [void](Invoke-ToolkitScript -ScriptPath (Join-Path $script:Root 'Reports/Get-M365UserPermissionsReport.ps1') -Parameters $parameters)
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

        Write-Host ''
        Write-Explain @(
            'Guests already in the tenant are not emailed by default - only ones',
            'created on this run. Say yes below to email them too, which is what',
            'you want if you added guests by hand and now need to tell them.',
            '',
            'It does not create a second account and does not disturb anyone''s',
            'group memberships: the existing guest is matched on their email',
            'address and only an invitation email is sent.'
        )

        if (Read-YesNo -Prompt '   Also email guests who already exist?' -Default $false) {
            $parameters['ResendInvitations'] = $true
        }
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
    $email = if ($parameters.ContainsKey('ResendInvitations')) { 'and email an invitation to every guest in the file, new or existing' }
             elseif ($parameters.ContainsKey('SendInvitationMessage')) { 'and email each new guest an invitation' }
             else { 'without emailing them' }

    Invoke-WithRehearsal -ScriptPath (Join-Path $script:Root 'Guests/Import-GuestPermissions.ps1') `
                         -Parameters $parameters `
                         -Description "create $count in the new tenant $email, and add them to their groups"
}

function Invoke-NotifyGuests {

    Write-Banner 'Email guests you have already created'

    Write-Explain @(
        'Sends the invitation email to guests who are already in the tenant -',
        'because you created them silently, or added them by hand - without',
        'changing anything else.',
        '',
        'No group membership is touched. Anyone given extra groups since they were',
        'migrated keeps them. No second account is created either: the existing',
        'guest is matched on their email address and only an email is sent.',
        '',
        'Everyone listed in the spreadsheet is emailed, so trim it first if some of',
        'them should not be told yet.'
    )

    Write-Step 'Which spreadsheet?'

    Write-Explain @('The same export you used to create them.')

    $parameters = @{
        InputPath             = Read-Text -Prompt '   Path to the exported .csv' -Default './TenantA_GuestPermissions.csv' `
            -Validate { param($v) Test-Path -Path $v } `
            -ValidationMessage 'No file at that path. Check the location and try again.'

        SendInvitationMessage = $true
        ResendInvitations     = $true
        SkipGroupMembership   = $true
    }

    Write-Step 'Have any of them already signed in?'

    Write-Explain @(
        'Microsoft only emails an invitation to a guest who has not accepted one yet.',
        'Once a guest has accepted, asking for another invitation is accepted without',
        'complaint but no email ever arrives. That is the usual reason for "I ran it',
        'and nobody got anything".',
        '',
        'Answering yes below resets those guests so the email is delivered. They keep',
        'their account, their groups and their app access, but the next time they open',
        'a resource they will be asked to accept the invitation again.',
        '',
        'Guests who have never accepted are emailed either way, so answer no if you',
        'are not sure - the log will tell you afterwards whether anyone was skipped.'
    )

    if (Read-YesNo -Prompt '   Re-invite guests who have already accepted') {
        $parameters['ResetRedemption'] = $true
    }

    Write-Step 'Anything to say to them?'

    $message = Read-Text -Prompt '   Message to include in the email (press Enter to skip)' -AllowEmpty

    if ($message) { $parameters['CustomInvitationMessage'] = $message }

    $parameters['LogPath'] = Get-OutputPath -Prompt 'Where should the results log go?' -Default './TenantB_GuestNotify_Log.csv'

    Write-Explain @(
        'A browser will open for you to sign in to the tenant the guests are in.',
        '',
        'This runs twice: first a rehearsal that sends nothing, then - once you',
        'answer y to the question that follows it - the real thing. Pressing Enter',
        'at that question means no, and nothing is sent.'
    )

    Invoke-WithRehearsal -ScriptPath (Join-Path $script:Root 'Guests/Import-GuestPermissions.ps1') `
                         -Parameters $parameters `
                         -Description 'email an invitation to every guest in the file, changing no group membership'

    Write-Explain @(
        'Check the Status column of the log:',
        '',
        '   InvitationResent   Microsoft accepted it and sent the email.',
        '   NoEmailSent        The guest had already accepted, so nothing was sent.',
        '                      Run this step again and answer yes to re-inviting them.',
        '   WhatIf             The rehearsal only. Nothing was sent - run it again and',
        '                      answer y when asked to go ahead.',
        '   Failed             Microsoft rejected it. The Detail column says why.',
        '',
        'If the log says InvitationResent and the guest still has nothing, the email',
        'left Microsoft and the delay is at their end - it comes from Microsoft',
        'Invitations, so junk mail and quarantine are the places to look.'
    )
}

function Invoke-SiteMembers {

    Write-Banner 'Report who can get into your SharePoint sites'

    Write-Explain @(
        'Produces a spreadsheet of everyone with access. Nothing is changed.',
        '',
        'Access reaches a site by several routes - the Members group, the Visitors',
        'group, any other group on the site, permissions given to somebody',
        'directly, and the connected Microsoft 365 group. All of them are listed,',
        'so a person who has access two ways appears twice. That is what you have',
        'to unpick to take their access away.',
        '',
        'Owners are reported by the other option, not this one.'
    )

    $parameters = Get-SiteSelection -AllowAllSites
    $parameters += Get-SharePointAuth

    Write-Step 'Who should be listed?'

    Write-Explain @(
        'Before a migration the useful question is usually which outside people',
        'can reach which sites. Answer yes to list only those.'
    )

    if (Read-YesNo -Prompt '   Guests only') { $parameters['GuestsOnly'] = $true }

    if (Read-YesNo -Prompt '   Include site owners as well') { $parameters['IncludeOwnersGroup'] = $true }

    Write-Step 'Where should the spreadsheet go?'

    $parameters['OutputPath'] = Get-OutputPath -Prompt 'File to create' -Default './SharePoint_SiteMembers.csv'

    Write-Step 'Running the report'

    Show-Command -ScriptPath (Join-Path $script:Root 'SharePoint/Get-SiteMembers.ps1') -Parameters $parameters

    [void](Invoke-ToolkitScript -ScriptPath (Join-Path $script:Root 'SharePoint/Get-SiteMembers.ps1') -Parameters $parameters)
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
        'and reduces any permission given to them directly on the site to read-only.',
        '',
        'Both matter: moving someone out of Members changes nothing if they were',
        'also given Edit directly on the site.',
        '',
        'People are added to Visitors before being removed from Members, so nobody',
        'is ever left without access.'
    )

    $parameters = Get-SiteSelection -AllowAllSites

    # Tenant-wide and destructive is a different order of risk from one site, so
    # it gets its own confirmation before any of the other questions are asked.
    if ($parameters.ContainsKey('AllSites')) {

        Write-Host ''
        Write-Explain @(
            'This will change permissions on EVERY site in the tenant, not just',
            'the ones you administer.',
            '',
            'Personal OneDrive sites are left out. Sites with no Visitors group are',
            'skipped rather than half-changed. You will still see a full preview',
            'before anything is applied.'
        )

        if (-not (Read-YesNo -Prompt '   Are you sure you want every site?' -Default $false)) {
            Write-Host ''
            Write-Host '   Cancelled. Nothing was changed.' -ForegroundColor Yellow
            return
        }
    }

    $parameters += Get-SharePointAuth

    Write-Step 'Who should be moved?'

    Write-Explain @('This applies to group membership and direct permissions alike.')

    $parameters['Scope'] = Read-Choice -Options @(
        @{ Key = 'Guests'; Label = 'Guests only'
           Detail = 'External people. Your own staff keep their current access.' }
        @{ Key = 'Staff'; Label = 'Staff only'
           Detail = 'Your own users. Guests keep their current access.' }
        @{ Key = 'Both'; Label = 'Guests and staff'
           Detail = 'Everyone with edit access to the site.' }
    )

    Write-Step 'Other groups on the site'

    Write-Explain @(
        'Some sites carry more than one member-type group - a plain "Members"',
        'alongside the site''s own "<Site> Members", left behind by whoever set it',
        'up. Emptying only the site''s own group leaves that other one exactly as',
        'it was, and the people in it keep editing.',
        '',
        'Answering yes empties every other group on the site into Visitors too.',
        'The Owners group and the Visitors group are never touched, whatever they',
        'have been renamed to.',
        '',
        'Worth knowing before you say yes: a custom group can carry Full Control,',
        'so read the preview - it names the group each person came from.'
    )

    if (Read-YesNo -Prompt '   Empty the site''s other groups as well' -Default $false) {
        $parameters['IncludeOtherGroups'] = $true
    }

    Write-Step 'Direct permissions'

    Write-Explain @(
        'As well as the Members group, people can be given access directly on a',
        'site. Those grants are reduced to read-only too, which is almost always',
        'what you want - otherwise someone moved out of Members carries on',
        'editing through their direct grant.',
        '',
        'Say yes here only if you want to change group membership and nothing else.'
    )

    if (Read-YesNo -Prompt '   Leave direct permissions alone?' -Default $false) {
        $parameters['SkipDirectPermissions'] = $true
    }

    Write-Step 'Teams sites'

    Write-Explain @(
        'On a site connected to a Team, the Members list contains the Team itself',
        '- shown as "<Site> Members" - as well as any individuals.',
        '',
        'That single entry is usually where most of the edit access comes from.',
        'Answer no and it stays exactly as it is: everyone in the Team carries on',
        'editing the site however many individuals were moved, which normally',
        'reads as the run having done nothing.',
        '',
        'Answer yes and the whole Team becomes read-only on this site in one go,',
        'including people who join later. It does not affect Teams chat, the',
        'shared mailbox or the calendar - Microsoft 365 group membership is never',
        'touched.',
        '',
        'The same answer covers any other security group in the Members list.'
    )

    if (Read-YesNo -Prompt '   Move the Team entry and any security groups as well' -Default $true) {
        $parameters['IncludeSecurityGroups'] = $true
    }

    Write-Explain @(
        'Moving that entry makes everyone in the Team read-only, but it does so',
        'through the Team - so none of them appear in the site''s Visitors list.',
        '',
        'Answer yes below to also add each member of the Team to Visitors by name.',
        'Nothing is removed and nobody leaves the Team; they are only added. Do',
        'this if the Team''s access to the site is going to change or go away, and',
        'these people still need to be able to read it.'
    )

    if (Read-YesNo -Prompt '   Also add each member of the Team to Visitors by name' -Default $false) {

        $parameters['AddGroupMembersAsVisitors'] = $true

        Write-Host ''
        Write-Explain @(
            'On a group-connected site, the Site permissions panel lists "Site',
            'members" from the Team itself, not from the SharePoint Members group.',
            'While somebody is in the Team they appear there whatever you do to the',
            'SharePoint groups, so the only way to clear them from that list is to',
            'take them out of the Team.',
            '',
            'That is not a SharePoint change. Leaving the Team also removes them',
            'from its chat and channels, the group mailbox, the group calendar and',
            'anything else attached to the group. They keep read access to the',
            'site, because they are added to Visitors first.',
            '',
            'Owners of the Team are never removed - they are reported and left.'
        )

        if (Read-YesNo -Prompt '   Remove them from the Team as well' -Default $false) {
            $parameters['RemoveFromMicrosoft365Group'] = $true
        }
    }

    Write-Step 'Anyone to leave alone?'

    Write-Explain @('For example an administrator or service account that must keep editing.')

    $exclude = Read-Text -Prompt '   Email addresses to skip, separated by commas (press Enter for none)' -AllowEmpty

    if ($exclude) {
        $parameters['ExcludeLogin'] = @($exclude -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    $parameters['OutputPath'] = Get-OutputPath -Prompt 'Where should the results log go?' -Default './SharePoint_MembersToViewers_Log.csv'

    $scope = switch ($parameters['Scope']) {
        'Both'  { 'guests and staff' }
        'Staff' { 'staff' }
        default { 'guests' }
    }

    $what = if ($parameters.ContainsKey('SkipDirectPermissions')) { ' in the Members group' }
            else { ' in the Members group and in direct site permissions' }

    $team  = if ($parameters.ContainsKey('IncludeSecurityGroups')) { ', including the connected Team' } else { '' }

    $where = if ($parameters.ContainsKey('AllSites'))      { ' on EVERY site in the tenant' }
             elseif ($parameters.ContainsKey('SitesCsvPath')) { ' on every site in your list' }
             else                                             { '' }

    Invoke-WithRehearsal -ScriptPath (Join-Path $script:Root 'SharePoint/Set-SiteMembersToViewers.ps1') `
                         -Parameters $parameters `
                         -Description "move $scope$what$team to view-only$where"
}

function Invoke-RegisterApp {

    Write-Banner 'Register the app in Entra ID'

    Write-Explain @(
        'The SharePoint tasks need an app registration in your tenant - Microsoft',
        'removed the shared one PnP used to rely on. This creates it and prints',
        'the client ID that every SharePoint task then asks for.',
        '',
        'Run this once per tenant, signed in as someone who can create app',
        'registrations and consent to permissions - Application Administrator or',
        'Global Administrator.'
    )

    $command = 'Register-PnPEntraIDAppForInteractiveLogin'

    if (-not (Get-Command -Name $command -ErrorAction SilentlyContinue)) {
        try { Import-Module PnP.PowerShell -ErrorAction Stop } catch { }
    }

    if (-not (Get-Command -Name $command -ErrorAction SilentlyContinue)) {
        Write-Host '   PnP.PowerShell is not available, so the app cannot be registered here.' -ForegroundColor Red
        Write-Host '   Run "Check and install what is needed" first.' -ForegroundColor Yellow
        Write-Host ''
        return
    }

    Write-Step 'Which tenant?'

    $tenant = Read-Text -Prompt '   Tenant name (e.g. contoso.onmicrosoft.com)' `
        -Validate ${function:Test-LooksLikeTenant} `
        -ValidationMessage 'That should look like contoso.onmicrosoft.com.'

    Write-Step 'What should it be called?'

    Write-Explain @('This is just a label, so you can find it in the Entra portal later.')

    $appName = Read-Text -Prompt '   Application name' -Default 'PnP Migration'

    $parameters = @{
        ApplicationName                = $appName
        Tenant                         = $tenant
        SharePointDelegatePermissions  = 'AllSites.FullControl'
        GraphDelegatePermissions       = 'Group.Read.All','User.Read.All'
    }

    Write-Step 'How should you sign in to create it?'

    $how = Read-Choice -Options @(
        @{ Key = 'browser'; Label = 'Open a browser'; Detail = 'Usual choice.' }
        @{ Key = 'device';  Label = 'Show me a code to enter on another device'
           Detail = 'For a machine with no browser, or when the browser flow is awkward with MFA.' }
    )

    if ($how -eq 'device') { $parameters['DeviceLogin'] = $true }

    Write-Step 'Creating the registration'

    Write-Explain @(
        'These permissions are what the scripts actually use: full control of',
        'SharePoint sites, and read-only on groups and users so the owner report',
        'can resolve names. Left unspecified, the command grants more than that.'
    )

    Show-Command -CommandName $command -Parameters $parameters

    try {
        $result = & $command @parameters

        Write-Host ''
        Write-Host '   Registered.' -ForegroundColor Green

        # The property name has varied between PnP releases, so take whichever is
        # present rather than assuming one.
        $clientId = @($result.'AzureAppId', $result.'AppId', $result.'ApplicationId') |
                        Where-Object { $_ } | Select-Object -First 1

        if ($clientId) {
            Write-Host ''
            Write-Host "   Client ID: $clientId" -ForegroundColor Green
            Write-Host '   Write this down - every SharePoint task asks for it.' -ForegroundColor Yellow
        }
        else {
            Write-Host '   Look for the client ID in the output above, and write it down.' -ForegroundColor Yellow
        }

        Write-Host ''
        Write-Host '   Consent can take a minute or two to apply. If the first SharePoint' -ForegroundColor DarkGray
        Write-Host '   task fails on permissions, wait and try again.' -ForegroundColor DarkGray
        Write-Host ''
    }
    catch {
        Write-Host ''
        Write-Host '   That did not work:' -ForegroundColor Red

        foreach ($line in ($_.Exception.Message -split "`n")) {
            Write-Host "     $($line.TrimEnd())" -ForegroundColor Red
        }

        Write-Host ''
        Write-Host '   Most often this means the account you signed in with cannot create' -ForegroundColor Yellow
        Write-Host '   app registrations, or cannot consent to permissions.' -ForegroundColor Yellow
        Write-Host ''
    }
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
        @{ Header = 'First-time setup'; Note = 'Run once on this machine.' }
        @{ Key = 'check'; Label = 'Check and install what is needed'; Tag = 'read-only'
           Detail = 'Start here. Lists what is missing and offers to install it.' }
        @{ Key = 'register'; Label = 'Register the app in Entra ID'; Tag = 'changes tenant'
           Detail = 'Needed before any SharePoint task. Gives you the client ID they ask for.' }
        @{ Key = 'cert';  Label = 'Create a certificate for full site access'; Tag = 'local files'
           Detail = 'Only needed to reach sites you do not administer yourself.' }

        @{ Header = 'Move guests to a new tenant'; Note = 'Do these in order - step 2 reads what step 1 writes.' }
        @{ Key = 'export'; Label = 'Step 1 - Export guests from the old tenant'; Tag = 'read-only'
           Detail = 'Writes a spreadsheet of guests and their groups.' }
        @{ Key = 'import'; Label = 'Step 2 - Create those guests in the new tenant'; Tag = 'changes tenant'
           Detail = 'Rehearsed first. Nothing is created until you confirm.' }
        @{ Key = 'notify'; Label = 'Step 3 - Email guests you have already created'; Tag = 'sends email'
           Detail = 'For guests created silently or added by hand. Changes no permissions.' }

        @{ Header = 'SharePoint permissions'; Note = 'Independent of the guest migration.' }
        @{ Key = 'owners';  Label = 'Report who owns your sites'; Tag = 'read-only'
           Detail = 'Writes a spreadsheet. Good to run before changing anything.' }
        @{ Key = 'members'; Label = 'Report who can get into your sites'; Tag = 'read-only'
           Detail = 'Members, visitors, other groups, direct permissions. Guests only, if you like.' }
        @{ Key = 'viewers'; Label = 'Change site members to view-only'; Tag = 'changes tenant'
           Detail = 'Groups and direct permissions. One site, a list, or the whole tenant.' }

        @{ Header = 'Reports'; Note = 'Point-in-time snapshots. Worth keeping as a record.' }
        @{ Key = 'allusers'; Label = 'List all users and their group access'; Tag = 'read-only'
           Detail = 'Staff and guests, with every group each one can reach.' }

        @{ Header = 'Finish' }
        @{ Key = 'quit'; Label = 'Quit' }
    )

    switch ($task) {
        'export'  { Invoke-ExportGuests }
        'import'  { Invoke-ImportGuests }
        'notify'  { Invoke-NotifyGuests }
        'owners'  { Invoke-SiteOwners }
        'members' { Invoke-SiteMembers }
        'viewers' { Invoke-MembersToViewers }
        'register' { Invoke-RegisterApp }
        'cert'    { Invoke-CreateCertificate }
        'allusers' { Invoke-UserPermissionsReport }
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
