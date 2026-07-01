<#
.SYNOPSIS
    Unzip a CSV dropped on disk and bulk-load it into a pre-existing table in the
    target on-premise SQL Server database, verifying the row count, then delete the zip.

.DESCRIPTION
    Called by Ansible during a deployment against an on-premise SQL Server 2019 host.
    Given a zip on D: (produced from a 20-column Excel -> CSV -> zip, bundled in an MSI),
    it:
      1. extracts the CSV,
      2. reads it into an in-memory table,
      3. verifies the target table and its columns exist (the table is PRE-CREATED with
         proper typed columns -- this script runs NO DDL),
      4. in ONE transaction: TRUNCATEs the table, bulk-copies the CSV rows, counts the
         table, and commits ONLY if the table count equals the CSV record count
         (mismatch -> ROLLBACK, so a bad load leaves the previous data intact), and
      5. on success, deletes the zip from the drive.

    Re-runnable / idempotent: TRUNCATE-then-load means the table always ends with exactly
    the CSV rows, so the count check passes on every redeploy.

    AUTHENTICATION: Windows / integrated (trusted connection) only -- there are no
    credential parameters and no secrets. The effective SQL identity is whatever account
    Ansible runs this script as (the WinRM user, or a `become` user); that account needs a
    SQL login plus INSERT and ALTER (ALTER is required for TRUNCATE) on the target table.

    RUNTIME: Windows PowerShell 5.1 (the default on Windows Server 2019), invoked as
    powershell.exe. It uses only built-in .NET Framework types
    (System.Data.SqlClient.SqlConnection / SqlBulkCopy, System.Data.DataTable) and built-in
    cmdlets (Expand-Archive, Import-Csv) -- NO module install (no SqlServer / Invoke-Sqlcmd)
    and no NuGet on the DB host. (Under PowerShell 7 it works only if a SqlClient provider is
    present; 5.1 is the supported path.)

    EXIT CODES:
      0  success (loaded + verified; zip deleted unless -KeepZip), or a successful -DryRun.
      1  general failure (unexpected error, SqlClient unavailable).
      2  invalid arguments, or row-count mismatch (load rolled back).
      3  missing prerequisite (zip/CSV not found, target table/column missing).

.PARAMETER ZipPath
    Path to the zip on disk (e.g. D:\deploy\data.zip). Required.

.PARAMETER SqlServer
    SQL Server instance. Default: localhost. Example: MYHOST\SQL2019.

.PARAMETER Database
    Target database. Default: database (override with your database name).

.PARAMETER Table
    Target table as schema.table (e.g. dbo.MyData) or a bare table (defaults schema to dbo).
    Validated against a safe identifier pattern before use (no injection via the name).

.PARAMETER ExtractDir
    Base directory to extract into. Default: the folder containing the zip. The archive is
    always expanded into a private "_csvload_<zipname>" subfolder that is cleaned each run.

.PARAMETER CsvName
    File name of the CSV to load when the archive contains more than one .csv.

.PARAMETER Delimiter
    CSV field delimiter. Default: comma. Use ';' for European-locale exports.

.PARAMETER Encoding
    CSV encoding passed to Import-Csv. Default: UTF8 (handles an Excel "CSV UTF-8" BOM).
    Use 'Default' (ANSI) for a plain Excel "CSV" export on a Western-European codepage.

.PARAMETER BatchSize
    SqlBulkCopy batch size. Default: 5000.

.PARAMETER CommandTimeout
    Per-command / bulk-copy timeout in seconds. Default: 300.

.PARAMETER UseDelete
    Use DELETE FROM instead of TRUNCATE TABLE (for a table referenced by a foreign key,
    where TRUNCATE is not permitted).

.PARAMETER KeepZip
    Do not delete the zip after a successful load (default is to delete it).

.PARAMETER KeepExtracted
    Do not delete the extracted files after a successful load.

.PARAMETER ReportPath
    Optional path for a small JSON run report.

.PARAMETER DryRun
    Extract and validate (table + columns + CSV count) but make NO changes: no truncate,
    no load, no delete.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\load-csv-to-sql.ps1 `
        -ZipPath D:\deploy\data.zip -SqlServer localhost -Database MyDatabase -Table dbo.MyData

.EXAMPLE
    # Validate only, no changes:
    powershell.exe -File .\load-csv-to-sql.ps1 -ZipPath D:\deploy\data.zip -Table dbo.MyData -DryRun
#>
[CmdletBinding()]
param(
    # ZipPath/Table are effectively required but NOT declared [Mandatory]: a mandatory param
    # would make PowerShell PROMPT when the value is missing, which hangs a non-interactive
    # Ansible call (and breaks dot-sourcing for tests). Invoke-Main validates them and returns 2.
    [string]$ZipPath,
    [string]$SqlServer = 'localhost',
    [string]$Database = 'database',
    [string]$Table,
    [string]$ExtractDir,
    [string]$CsvName,
    [string]$Delimiter = ',',
    [string]$Encoding = 'UTF8',
    [int]$BatchSize = 5000,
    [int]$CommandTimeout = 300,
    [switch]$UseDelete,
    [switch]$KeepZip,
    [switch]$KeepExtracted,
    [string]$ReportPath,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'   # cmdlet failures (Expand-Archive, Import-Csv, SQL calls) throw

# --- pure helpers (dot-sourceable; unit-tested by tests/load-csv-to-sql.Tests.ps1) ---

function Test-SqlIdentifier {
    # Accept 'table' or 'schema.table'. Each part: a letter/underscore, then word / @ $ # chars.
    # No spaces, brackets, quotes or semicolons -> the name is safe to wrap in [] for SQL.
    param([Parameter(Mandatory = $true)][string]$Name)
    return $Name -match '^[A-Za-z_][A-Za-z0-9_@$#]*(\.[A-Za-z_][A-Za-z0-9_@$#]*)?$'
}

function Split-SqlTable {
    # 'dbo.MyData' -> @{Schema='dbo';Table='MyData'};  'MyData' -> @{Schema='dbo';Table='MyData'}
    param([Parameter(Mandatory = $true)][string]$Name)
    $bits = $Name.Split('.')
    if ($bits.Count -eq 2) { return @{ Schema = $bits[0]; Table = $bits[1] } }
    return @{ Schema = 'dbo'; Table = $bits[0] }
}

function Get-QuotedTableName {
    # Build [schema].[table] from already-validated parts (safe: Test-SqlIdentifier forbids ]).
    param([Parameter(Mandatory = $true)][hashtable]$Parts)
    return "[$($Parts.Schema)].[$($Parts.Table)]"
}

function Format-ConnValue {
    # Quote a connection-string value only if it contains a special char or edge whitespace.
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Value)
    if ($Value -match '[";=]' -or $Value -ne $Value.Trim()) {
        if ($Value.Contains('"')) { return "'" + $Value + "'" }
        return '"' + $Value + '"'
    }
    return $Value
}

function New-SqlConnectionString {
    # Windows integrated auth only -- no user, no password. SqlClient-free (buildable anywhere).
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$Database
    )
    $ds = Format-ConnValue $Server
    $ic = Format-ConnValue $Database
    return "Data Source=$ds;Initial Catalog=$ic;Integrated Security=SSPI;Application Name=load-csv-to-sql;TrustServerCertificate=True"
}

function Get-CsvColumns {
    # Header names, in order. From the first record when there is data; else parse the header
    # line (best-effort split -- only used for a header-only CSV with zero data rows).
    param([object[]]$Records, [string]$HeaderLine, [string]$Delimiter = ',')
    if ($Records -and $Records.Count -gt 0) {
        return @($Records[0].PSObject.Properties.Name)
    }
    if (-not [string]::IsNullOrWhiteSpace($HeaderLine)) {
        $line = $HeaderLine.TrimStart([char]0xFEFF)
        return @($line.Split($Delimiter) | ForEach-Object { $_.Trim().Trim('"') })
    }
    return @()
}

function Get-MissingColumns {
    # CSV columns absent from the table (case-insensitive) -> a non-empty result is a mismatch.
    param([string[]]$CsvColumns, [string[]]$TableColumns)
    $set = @{}
    foreach ($t in $TableColumns) { $set[$t.ToLowerInvariant()] = $true }
    return @($CsvColumns | Where-Object { -not $set.ContainsKey($_.ToLowerInvariant()) })
}

function Resolve-CsvFile {
    # Pick the CSV from a list of extracted file paths (optionally by name).
    param([string[]]$Files, [string]$CsvName)
    $csvs = @($Files | Where-Object { [IO.Path]::GetExtension($_) -ieq '.csv' })
    if (-not [string]::IsNullOrWhiteSpace($CsvName)) {
        $match = @($csvs | Where-Object { [IO.Path]::GetFileName($_) -ieq $CsvName })
        if ($match.Count -eq 0) { throw "CSV '$CsvName' not found in the archive." }
        return $match[0]
    }
    if ($csvs.Count -eq 0) { throw "No .csv file found in the archive." }
    if ($csvs.Count -gt 1) {
        throw "Multiple CSV files found; pass -CsvName to choose one of: $(($csvs | ForEach-Object { [IO.Path]::GetFileName($_) }) -join ', ')"
    }
    return $csvs[0]
}

function ConvertTo-DataTable {
    # CSV records -> a string-typed DataTable. Empty strings become DBNull so nullable typed
    # destination columns receive NULL rather than failing conversion.
    param(
        [object[]]$Records,
        [Parameter(Mandatory = $true)][string[]]$Columns
    )
    $dt = New-Object System.Data.DataTable
    foreach ($c in $Columns) { [void]$dt.Columns.Add($c, [string]) }
    foreach ($rec in $Records) {
        $row = $dt.NewRow()
        foreach ($c in $Columns) {
            $val = $rec.$c
            if ($null -eq $val -or ($val -is [string] -and $val -eq '')) {
                $row[$c] = [System.DBNull]::Value
            }
            else {
                $row[$c] = [string]$val
            }
        }
        [void]$dt.Rows.Add($row)
    }
    return , $dt   # unary comma: return the DataTable itself, not its enumerated rows
}

function New-LoadReport {
    param(
        [string]$ZipPath, [string]$Csv, [string]$Table,
        [int64]$CsvCount, [int64]$TableCount, [string]$Status,
        [bool]$ZipDeleted, [bool]$DryRun
    )
    return [ordered]@{
        zip_path    = $ZipPath
        csv         = $Csv
        table       = $Table
        csv_count   = $CsvCount
        table_count = $TableCount
        status      = $Status
        zip_deleted = $ZipDeleted
        dry_run     = $DryRun
    }
}

# --- SQL steps (SqlClient; resolved at call time, so dot-sourcing never needs the provider) ---

function Get-TableColumns {
    # Column names of Schema.Table via INFORMATION_SCHEMA (parameterised). Empty => table absent.
    param($Connection, [string]$Schema, [string]$Table)
    $cmd = $Connection.CreateCommand()
    $cmd.CommandText = 'SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = @s AND TABLE_NAME = @t'
    [void]$cmd.Parameters.AddWithValue('@s', $Schema)
    [void]$cmd.Parameters.AddWithValue('@t', $Table)
    $cols = [System.Collections.ArrayList]::new()
    $reader = $cmd.ExecuteReader()
    try { while ($reader.Read()) { [void]$cols.Add([string]$reader['COLUMN_NAME']) } }
    finally { $reader.Close() }
    return $cols.ToArray()
}

function Invoke-Load {
    # One transaction: clear -> bulk-copy -> count -> commit iff table count == CSV count.
    param(
        $Connection, $DataTable, [string]$QuotedTable, [string[]]$Columns,
        [int]$BatchSize, [int]$CommandTimeout, [bool]$UseDelete
    )
    $tx = $Connection.BeginTransaction('csvload')
    try {
        $clearSql = if ($UseDelete) { "DELETE FROM $QuotedTable;" } else { "TRUNCATE TABLE $QuotedTable;" }
        $clear = $Connection.CreateCommand()
        $clear.Transaction = $tx
        $clear.CommandTimeout = $CommandTimeout
        $clear.CommandText = $clearSql
        [void]$clear.ExecuteNonQuery()

        $opts = [System.Data.SqlClient.SqlBulkCopyOptions]::TableLock
        $bulk = New-Object System.Data.SqlClient.SqlBulkCopy($Connection, $opts, $tx)
        try {
            $bulk.DestinationTableName = $QuotedTable
            $bulk.BatchSize = $BatchSize
            $bulk.BulkCopyTimeout = $CommandTimeout
            foreach ($c in $Columns) { [void]$bulk.ColumnMappings.Add($c, $c) }
            $bulk.WriteToServer($DataTable)
        }
        finally { $bulk.Close() }

        $countCmd = $Connection.CreateCommand()
        $countCmd.Transaction = $tx
        $countCmd.CommandTimeout = $CommandTimeout
        $countCmd.CommandText = "SELECT COUNT_BIG(*) FROM $QuotedTable;"
        $tableCount = [int64]$countCmd.ExecuteScalar()
        $csvCount = [int64]$DataTable.Rows.Count

        if ($tableCount -ne $csvCount) {
            $tx.Rollback()
            return [pscustomobject]@{ Committed = $false; CsvCount = $csvCount; TableCount = $tableCount }
        }
        $tx.Commit()
        return [pscustomobject]@{ Committed = $true; CsvCount = $csvCount; TableCount = $tableCount }
    }
    catch {
        try { $tx.Rollback() } catch { }
        throw
    }
}

# --- side-effecting helpers ---

function Write-Err {
    # Controlled-failure diagnostic to stderr. Deliberately NOT Write-Error: with
    # $ErrorActionPreference='Stop' that throws, which would skip the specific `return <code>`
    # and collapse every controlled exit to 1. Ansible captures stderr + the exit code.
    param([Parameter(Mandatory = $true)][string]$Message)
    [Console]::Error.WriteLine("ERROR: $Message")
}

function Write-LoadReport {
    param([object]$Report)
    if ([string]::IsNullOrWhiteSpace($ReportPath)) { return }
    $dir = Split-Path -Parent $ReportPath
    if ($dir) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $Report | ConvertTo-Json -Depth 5 | Out-File -FilePath $ReportPath -Encoding utf8
    Write-Host "Report: $ReportPath"
}

# --- orchestration ---

function Invoke-Main {
    # 1. validate arguments
    if ([string]::IsNullOrWhiteSpace($ZipPath)) { Write-Err '-ZipPath is required.'; return 2 }
    if ([string]::IsNullOrWhiteSpace($Table)) { Write-Err '-Table is required.'; return 2 }
    if (-not (Test-SqlIdentifier -Name $Table)) {
        Write-Err "-Table '$Table' is not a valid identifier (expected schema.table or table; letters, digits and _ @ $ # only)."
        return 2
    }
    if (-not (Test-Path -LiteralPath $ZipPath)) { Write-Err "Zip not found: $ZipPath"; return 3 }
    if (-not ('System.Data.SqlClient.SqlConnection' -as [type])) {
        Write-Err 'System.Data.SqlClient is unavailable. Run with Windows PowerShell 5.1 (powershell.exe), or install a SqlClient provider for PowerShell 7.'
        return 1
    }

    $parts = Split-SqlTable -Name $Table
    $quoted = Get-QuotedTableName -Parts $parts

    # 2. extract into a private, per-zip subfolder (cleaned each run)
    $zipFull = (Resolve-Path -LiteralPath $ZipPath).Path
    $base = if (-not [string]::IsNullOrWhiteSpace($ExtractDir)) { $ExtractDir } else { Split-Path -Parent $zipFull }
    $work = Join-Path $base ('_csvload_' + [IO.Path]::GetFileNameWithoutExtension($zipFull))
    if (Test-Path -LiteralPath $work) { Remove-Item -LiteralPath $work -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $work | Out-Null
    Write-Host "Extract: $zipFull -> $work"
    Expand-Archive -LiteralPath $zipFull -DestinationPath $work -Force

    $files = @(Get-ChildItem -LiteralPath $work -Recurse -File | Select-Object -ExpandProperty FullName)
    try { $csvPath = Resolve-CsvFile -Files $files -CsvName $CsvName }
    catch { Write-Err $_.Exception.Message; return 3 }
    Write-Host "CSV: $csvPath"

    # 3. read CSV -> DataTable
    $records = @(Import-Csv -LiteralPath $csvPath -Delimiter $Delimiter -Encoding $Encoding)
    $headerLine = (Get-Content -LiteralPath $csvPath -TotalCount 1)
    $columns = Get-CsvColumns -Records $records -HeaderLine $headerLine -Delimiter $Delimiter
    if ($columns.Count -eq 0) { Write-Err "Could not determine CSV columns from $csvPath."; return 3 }
    $dt = ConvertTo-DataTable -Records $records -Columns $columns
    $csvCount = [int64]$dt.Rows.Count
    Write-Host "CSV rows: $csvCount ; columns ($($columns.Count)): $($columns -join ', ')"

    # 4. connect + pre-checks (before any mutation)
    $conn = New-Object System.Data.SqlClient.SqlConnection((New-SqlConnectionString -Server $SqlServer -Database $Database))
    $conn.Open()
    try {
        $tableCols = Get-TableColumns -Connection $conn -Schema $parts.Schema -Table $parts.Table
        if ($tableCols.Count -eq 0) {
            Write-Err "Table $quoted not found in [$Database] (schema '$($parts.Schema)'). Create it before deploying."
            return 3
        }
        $missing = Get-MissingColumns -CsvColumns $columns -TableColumns $tableCols
        if ($missing.Count -gt 0) {
            Write-Err "CSV column(s) not present in $quoted : $($missing -join ', ')"
            return 3
        }

        if ($DryRun) {
            $clearVerb = if ($UseDelete) { 'DELETE FROM' } else { 'TRUNCATE' }
            Write-Host "DRYRUN  would $clearVerb $quoted and load $csvCount row(s). No changes made."
            Write-LoadReport (New-LoadReport -ZipPath $zipFull -Csv $csvPath -Table $quoted -CsvCount $csvCount -TableCount ([int64]-1) -Status 'dryrun' -ZipDeleted $false -DryRun $true)
            return 0
        }

        # 5. transactional truncate + load + count
        $res = Invoke-Load -Connection $conn -DataTable $dt -QuotedTable $quoted -Columns $columns `
            -BatchSize $BatchSize -CommandTimeout $CommandTimeout -UseDelete:([bool]$UseDelete)
        if (-not $res.Committed) {
            Write-Err "Row-count mismatch: CSV=$($res.CsvCount) table=$($res.TableCount). Rolled back; table unchanged."
            Write-LoadReport (New-LoadReport -ZipPath $zipFull -Csv $csvPath -Table $quoted -CsvCount $res.CsvCount -TableCount $res.TableCount -Status 'count-mismatch' -ZipDeleted $false -DryRun $false)
            return 2
        }
        Write-Host "Loaded and verified: CSV=$($res.CsvCount) == table=$($res.TableCount)."
    }
    finally {
        $conn.Close()
        $conn.Dispose()
    }

    # 6. cleanup on success (a failed cleanup is a warning, not a failure -- the load committed)
    $zipDeleted = $false
    if (-not $KeepZip) {
        try { Remove-Item -LiteralPath $zipFull -Force; $zipDeleted = $true; Write-Host "Deleted zip: $zipFull" }
        catch { Write-Warning "Load succeeded but the zip could not be deleted: $zipFull -- $($_.Exception.Message)" }
    }
    if (-not $KeepExtracted) {
        try { Remove-Item -LiteralPath $work -Recurse -Force }
        catch { Write-Warning "Could not remove the extract folder $work : $($_.Exception.Message)" }
    }

    # 7. report
    Write-LoadReport (New-LoadReport -ZipPath $zipFull -Csv $csvPath -Table $quoted -CsvCount $csvCount -TableCount $csvCount -Status 'ok' -ZipDeleted $zipDeleted -DryRun $false)
    return 0
}

# Run only when invoked directly (powershell.exe -File / & script.ps1), NOT when dot-sourced.
if ($MyInvocation.InvocationName -ne '.') {
    try {
        exit (Invoke-Main)
    }
    catch {
        # Any unexpected error (bad zip, connection/permission failure, bulk-copy error) ->
        # clean stderr line + exit 1. The zip is left in place (cleanup runs only on success).
        [Console]::Error.WriteLine("ERROR: $($_.Exception.Message)")
        exit 1
    }
}
