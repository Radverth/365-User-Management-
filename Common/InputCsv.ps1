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
        [string[]]$RequiredColumns = @()
    )

    Assert-TextCsv -Path $Path

    $candidates = if ($Delimiter) { @($Delimiter) } else { @(',', ';', "`t", '|') }

    $best = $null

    foreach ($candidate in $candidates) {

        try   { $rows = @(Import-Csv -Path $Path -Delimiter $candidate -ErrorAction Stop) }
        catch { continue }

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

    if ($null -eq $best) {
        throw "Could not read any rows from '$Path'. The file is empty, or contains only a header."
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
            [void]$message.AppendLine('  Pass the CSV produced by Export-GuestPermissions.ps1, not the unsupported-')
            [void]$message.AppendLine('  groups CSV or a report from another script.')
        }

        throw $message.ToString()
    }

    return ,$rows
}
