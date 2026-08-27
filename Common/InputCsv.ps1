<#
.SYNOPSIS
    Shared CSV input handling for the migration scripts.

.DESCRIPTION
    Dot-sourced by the scripts in Guests/ and SharePoint/. Not run directly.

    Import-Csv is unforgiving about how a file reached it, and in practice input
    files have usually been through Excel. This wraps it with the two checks that
    turn a baffling "missing required column" error into an actionable one:

      1. A binary spreadsheet (.xlsx / .xls) wearing a .csv extension is rejected
         with an explanation, instead of parsing as one nonsense column.
      2. The field separator is detected, so a file saved in a locale that uses
         semicolons still loads. A byte-order mark stuck to the first column name
         is trimmed off.

    Failures name the file, the columns that were actually found, and the likely
    fix.
#>

function Assert-TextCsv {
    <#  Rejects binary spreadsheets up front.

        Opening a CSV in Excel and using Save As easily produces a real .xlsx that
        keeps its .csv name. Import-Csv then reads the ZIP bytes as text, finds no
        delimiter on the first line and returns a single column, so every required
        column appears to be missing - which says nothing about the real cause. #>
    param([Parameter(Mandatory)][string]$Path)

    $header = [byte[]]::new(8)

    try {
        $stream = [System.IO.File]::OpenRead($Path)

        try   { $read = $stream.Read($header, 0, 8) }
        finally { $stream.Dispose() }
    }
    catch {
        return  # Unreadable files fail later with a clearer message.
    }

    if ($read -lt 4) { return }

    # PK\x03\x04 - ZIP container, so .xlsx / .xlsm / .ods
    $isZip = ($header[0] -eq 0x50 -and $header[1] -eq 0x4B -and $header[2] -eq 0x03 -and $header[3] -eq 0x04)

    # D0 CF 11 E0 - OLE2 compound file, so a legacy .xls
    $isOle = ($header[0] -eq 0xD0 -and $header[1] -eq 0xCF -and $header[2] -eq 0x11 -and $header[3] -eq 0xE0)

    if (-not ($isZip -or $isOle)) { return }

    $kind = if ($isZip) { 'an Excel workbook (.xlsx) or OpenDocument spreadsheet' } else { 'a legacy Excel workbook (.xls)' }

    $message = [System.Text.StringBuilder]::new()

    [void]$message.AppendLine("'$Path' is $kind, not a text CSV, despite its file name.")
    [void]$message.AppendLine('')
    [void]$message.AppendLine('  This happens when a CSV is opened in Excel and saved with File > Save As,')
    [void]$message.AppendLine('  which writes a workbook while leaving the .csv extension in place.')
    [void]$message.AppendLine('')
    [void]$message.AppendLine('  Fix it either way:')
    [void]$message.AppendLine('    - In Excel: File > Save As > CSV UTF-8 (Comma delimited) (*.csv)')
    [void]$message.AppendLine('    - Or re-run Export-GuestPermissions.ps1 and pass the file straight to this')
    [void]$message.AppendLine('      script without opening it in Excel.')

    throw $message.ToString()
}

function Assert-WrittenCsv {
    <#  Reads back a file we just wrote and checks it parses into the rows and
        columns that went in.

        Export-Csv is not supposed to be able to produce an unreadable file, but a
        report that cannot be read is worthless, and finding out at import time -
        possibly on a different machine, days later - wastes the whole round trip.
        Verifying at the point of writing turns that into an immediate, local
        failure with the evidence still to hand.

        Returns $true when the file reads back correctly. #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][int]$ExpectedRows,
        [string[]]$RequiredColumns = @()
    )

    if (-not (Test-Path -Path $Path)) {
        Write-Warning "The report was not written to $Path."
        return $false
    }

    $bytes = (Get-Item -Path $Path).Length
    $raw   = [System.IO.File]::ReadAllText($Path)
    $lines = ([regex]::Matches($raw, "`r`n|`n|`r")).Count

    $problems = [System.Collections.Generic.List[string]]::new()

    try {
        $readBack = @(Import-Csv -Path $Path -ErrorAction Stop)

        $names = if ($readBack.Count -gt 0) { @($readBack[0].PSObject.Properties.Name) } else { @() }

        if ($readBack.Count -ne $ExpectedRows) {
            $problems.Add("wrote $ExpectedRows row(s) but read back $($readBack.Count)")
        }

        foreach ($column in $RequiredColumns) {
            if ($column -notin $names) { $problems.Add("column '$column' is missing when read back") }
        }
    }
    catch {
        $problems.Add("reading it back failed: $($_.Exception.Message)")
    }

    if ($lines -lt $ExpectedRows) {
        $problems.Add("the file holds $lines line break(s) for $ExpectedRows row(s) plus a header")
    }

    if ($problems.Count -eq 0) { return $true }

    Write-Warning "The report at $Path did not read back correctly:"

    foreach ($problem in $problems) { Write-Warning "  - $problem" }

    Write-Warning "  File size: $bytes bytes, line breaks: $lines"
    Write-Warning '  Do not rely on this file. Please report it with the two numbers above.'

    return $false
}

function Import-InputCsv {
    <#  Import-Csv with the two failure modes that actually bite in the field.

        Wrong delimiter: a file re-saved by Excel under a locale that uses ';' (or
        exported as tab-separated) parses as ONE column holding the whole header
        line, so every required column looks missing. Each candidate delimiter is
        tried and scored on how many required columns it yields.

        Byte-order mark: a leading BOM can end up glued to the first column name,
        so '<BOM>GuestDisplayName' never matches 'GuestDisplayName'. Names are
        trimmed and the rows rebuilt when anything needed cleaning. #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Delimiter,
        [string[]]$RequiredColumns = @(),

        # What the caller wanted, named in the error when the wrong file arrives.
        # Without it the advice would have to be generic, or - worse - specific to
        # whichever caller was written first.
        [string]$Expected = 'the file this script expects'
    )

    Assert-TextCsv -Path $Path

    # Kept for diagnostics: a file whose records all sit on one line cannot be a
    # delimiter problem, and saying so beats guessing.
    $raw     = [System.IO.File]::ReadAllText($Path)
    $hasBreak = $raw -match "[`r`n]"

    $candidates = if ($Delimiter) { @($Delimiter) } else { @(',', ';', "`t", '|') }

    $best        = $null
    $parseErrors = @{}

    foreach ($candidate in $candidates) {

        try {
            $rows = @(Import-Csv -Path $Path -Delimiter $candidate -ErrorAction Stop)
        }
        catch {
            # Do NOT discard this. Import-Csv throws "The member ... is already
            # present" when the header carries a repeated name, which is what a
            # file with its records run together looks like. Swallowing it and
            # moving to the next delimiter reports a wrong cause with confidence.
            $parseErrors[$candidate] = $_.Exception.Message
            continue
        }

        if ($rows.Count -eq 0) { continue }

        $names   = @($rows[0].PSObject.Properties.Name)
        $trimmed = @($names | ForEach-Object { $_.Trim([char]0xFEFF, ' ', '"') })
        $matched = @($RequiredColumns | Where-Object { $_ -in $trimmed }).Count

        $result = [PSCustomObject]@{
            Delimiter = $candidate
            Rows      = $rows
            Names     = $names
            Trimmed   = $trimmed
            Matched   = $matched
        }

        if ($null -eq $best -or
            $matched -gt $best.Matched -or
            ($matched -eq $best.Matched -and $trimmed.Count -gt $best.Trimmed.Count)) {
            $best = $result
        }

        if ($RequiredColumns.Count -gt 0 -and $matched -eq $RequiredColumns.Count) { break }
    }

    # A single column means nothing was really split, so this is not a usable
    # parse - report the underlying cause rather than the delimiter that "won".
    $usable = $best -and ($best.Trimmed.Count -gt 1 -or $best.Matched -gt 0)

    if (-not $usable) {

        $message = [System.Text.StringBuilder]::new()

        [void]$message.AppendLine("Could not read '$Path' as a CSV.")
        [void]$message.AppendLine('')

        if (-not $hasBreak) {
            [void]$message.AppendLine('  The file contains no line breaks at all, so the header and every record')
            [void]$message.AppendLine('  are run together on a single line. Nothing can be read from it.')
            [void]$message.AppendLine('')
            [void]$message.AppendLine('  This usually means the file was converted or copied by something that')
            [void]$message.AppendLine('  stripped the line endings. Re-run the export to produce a fresh file,')
            [void]$message.AppendLine('  and copy it as a file rather than pasting its contents.')
        }
        elseif ($parseErrors.Count -gt 0) {

            foreach ($key in $parseErrors.Keys) {
                $shown = if ($key -eq "`t") { 'tab' } else { "'$key'" }
                [void]$message.AppendLine("  Reading it with $shown as the separator failed: $($parseErrors[$key])")
            }

            [void]$message.AppendLine('')
            [void]$message.AppendLine('  "The member ... is already present" means the header has the same column')
            [void]$message.AppendLine('  name twice, which happens when records have been run together onto one')
            [void]$message.AppendLine('  line. Re-run the export to produce a fresh file.')
        }
        else {
            [void]$message.AppendLine('  No separator produced more than a single column.')
            [void]$message.AppendLine('  Re-run the export, or pass the separator explicitly with -Delimiter.')
        }

        throw $message.ToString()
    }

    if ($best.Delimiter -ne ',' -and -not $Delimiter) {
        $shown = if ($best.Delimiter -eq "`t") { 'tab' } else { "'$($best.Delimiter)'" }
        Write-Host "  Detected $shown as the field separator." -ForegroundColor Yellow
    }

    $rows = $best.Rows

    # Rebuild only when a name actually needed cleaning, so the common path is free.
    $needsRebuild = $false

    for ($i = 0; $i -lt $best.Names.Count; $i++) {
        if ($best.Names[$i] -cne $best.Trimmed[$i]) { $needsRebuild = $true; break }
    }

    if ($needsRebuild) {
        $rows = @(
            foreach ($row in $rows) {
                $clean = [ordered]@{}

                for ($i = 0; $i -lt $best.Names.Count; $i++) {
                    $clean[$best.Trimmed[$i]] = $row.($best.Names[$i])
                }

                [PSCustomObject]$clean
            }
        )
    }

    $missing = @($RequiredColumns | Where-Object { $_ -notin $best.Trimmed })

    if ($missing.Count -gt 0) {

        $message = [System.Text.StringBuilder]::new()

        [void]$message.AppendLine("Input CSV is missing required column(s): $($missing -join ', ').")
        [void]$message.AppendLine("  File    : $Path")
        [void]$message.AppendLine("  Found   : $($best.Trimmed -join ', ')")

        if ($best.Trimmed.Count -eq 1) {
            # One column holding the whole header means the delimiter is wrong.
            [void]$message.AppendLine('')
            [void]$message.AppendLine('  The entire header parsed as a single column, so the field separator is wrong.')
            [void]$message.AppendLine('  This happens when the file has been opened and re-saved by Excel in a locale')
            [void]$message.AppendLine('  that uses semicolons. Re-export it, or pass the separator explicitly:')
            [void]$message.AppendLine('      -Delimiter ";"')
        }
        else {
            [void]$message.AppendLine('')
            [void]$message.AppendLine("  This does not look like $Expected.")
            [void]$message.AppendLine('  Check the Found line above against the columns listed as missing.')
        }

        throw $message.ToString()
    }

    return ,$rows
}
