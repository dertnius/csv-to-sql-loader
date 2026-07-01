# CSV → SQL Server loader (Ansible deployment step)

[![tests](https://github.com/dertnius/csv-to-sql-loader/actions/workflows/pester.yml/badge.svg)](https://github.com/dertnius/csv-to-sql-loader/actions/workflows/pester.yml)

`load-csv-to-sql.ps1` unzips a CSV that a deployment dropped on disk, bulk-loads it into a
**pre-created** table in the target on-premise SQL Server database, verifies the loaded row
count matches the CSV, and then deletes the zip. It is meant to be called by **Ansible** during a
deployment against a SQL Server 2019 host.

It is self-contained and independent of the rest of this repository — copy the `deploy/` folder
into your MSI/Ansible project wherever it is convenient.

## What it does

1. Extracts the CSV from the zip into a private, per-zip subfolder (cleaned each run).
2. Reads the CSV into memory (empty cells become `NULL`).
3. Validates the target table **and** every CSV column exist (no DDL — the table is pre-created).
4. In **one transaction**: `TRUNCATE` the table → `SqlBulkCopy` the rows → `SELECT COUNT(*)` →
   **commit only if the table count equals the CSV record count** (mismatch → rollback, so the
   previous data is left intact).
5. On success, deletes the zip.

Re-runnable / **idempotent**: because it truncates then loads, the table always ends with exactly
the CSV rows, so the count check passes on every redeploy.

## Prerequisites

- **Windows PowerShell 5.1** (the default on Windows Server 2019) — invoke as `powershell.exe`.
  No modules to install: it uses only built-in .NET (`System.Data.SqlClient`) and built-in cmdlets
  (`Expand-Archive`, `Import-Csv`). No `SqlServer` module, no `Invoke-Sqlcmd`, no NuGet.
- **Windows / integrated authentication.** There are no credential parameters and no secrets. The
  effective SQL identity is **the account that runs the script** (the Ansible WinRM user, or a
  `become` user). That account needs, on the target instance:
  - a **SQL login** with access to the target database, and
  - **`INSERT`** and **`ALTER`** on the target table (`ALTER` is required for `TRUNCATE`;
    membership in `db_ddladmin`/`db_owner` covers it). Use `-UseDelete` if you can only grant
    `DELETE` (e.g. the table is referenced by a foreign key).
- The target **table is pre-created** with proper typed columns. Any table columns **not** present
  in the CSV must be nullable, have a default, or be an `IDENTITY` (the loader does not populate
  them).

## Ansible usage

```yaml
- name: Load reference CSV into the database
  ansible.windows.win_command: >-
    powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass
    -File D:\deploy\load-csv-to-sql.ps1
    -ZipPath D:\deploy\data.zip
    -SqlServer localhost
    -Database MyDatabase
    -Table dbo.MyData
  register: csv_load
  failed_when: csv_load.rc != 0
```

- If the SQL host is a **different** machine from the WinRM target, integrated auth needs a
  double-hop — use **Kerberos** (or CredSSP) for the Ansible connection, otherwise the login will
  fail as `NT AUTHORITY\ANONYMOUS`.
- `win_command` does not run through a shell, so no quoting of a shell is involved; pass each flag
  as shown. (`ansible.windows.win_powershell` is an alternative if you prefer.)

## Parameters

| Parameter | Default | Notes |
|---|---|---|
| `-ZipPath` | *(required)* | Path to the zip on disk, e.g. `D:\deploy\data.zip`. |
| `-SqlServer` | `localhost` | Instance, e.g. `MYHOST\SQL2019`. |
| `-Database` | `database` | Target database. |
| `-Table` | *(required)* | `schema.table` (e.g. `dbo.MyData`) or a bare table (schema defaults to `dbo`). Validated against a safe identifier pattern. |
| `-ExtractDir` | folder of the zip | Base dir to extract into (a private `_csvload_<zip>` subfolder is used). |
| `-CsvName` | *(auto)* | Pick the CSV when the archive holds more than one `.csv`. |
| `-Delimiter` | `,` | Use `;` for European-locale exports. |
| `-Encoding` | `UTF8` | `UTF8` handles an Excel *"CSV UTF-8"* BOM; use `Default` (ANSI) for a plain Excel *"CSV"* on a Western-European codepage. |
| `-BatchSize` | `5000` | `SqlBulkCopy` batch size. |
| `-CommandTimeout` | `300` | Per-command / bulk-copy timeout (seconds). |
| `-UseDelete` | off | `DELETE FROM` instead of `TRUNCATE` (FK-referenced tables). |
| `-KeepZip` | off | Do not delete the zip after a successful load. |
| `-KeepExtracted` | off | Do not delete the extracted files after a successful load. |
| `-ReportPath` | *(none)* | Write a small JSON run report to this path. |
| `-DryRun` | off | Extract + validate only — no truncate, no load, no delete. |

## Exit codes

| Code | Meaning |
|---|---|
| `0` | Success (loaded + verified; zip deleted unless `-KeepZip`), or a successful `-DryRun`. |
| `1` | General failure (unexpected error; SQL connection/permission failure; bulk-copy error; SqlClient unavailable). |
| `2` | Invalid arguments, **or** row-count mismatch — the load was rolled back, so the table is unchanged. |
| `3` | Missing prerequisite: zip or CSV not found, or the target table/column does not exist. |

Any non-zero code fails the Ansible task. On any failure the **zip is left in place** for
troubleshooting.

## CSV / data caveats

- **Encoding & delimiter** must match the Excel export (see `-Encoding` / `-Delimiter` above). A
  wrong encoding shows up as garbled text or a leading BOM on the first column name.
- **Dates and numbers** are bulk-copied from text into the pre-created typed columns. Use an
  **unambiguous, culture-invariant** representation in the CSV — ISO-8601 dates (`2026-01-31`) and
  `.`-decimal numbers — so `SqlBulkCopy`'s string→type conversion is deterministic. A locale date
  like `31/01/2026` may fail to convert or land on the wrong day.
- **Empty cells** are loaded as `NULL`. A pre-created column that is `NOT NULL` with no default
  will therefore reject an empty cell (by design — fix the data or the schema).
- Column matching is **by name, case-insensitive**, and order-independent; a CSV column that has no
  matching table column stops the run before any data is touched (exit 3).

## Local testing

```powershell
# From this folder. Unit tests for the pure helpers (no SQL Server needed):
powershell.exe -NoProfile -Command "Invoke-Pester -Path .\tests -CI"

# Dry run against a real instance (validates table/columns/count, changes nothing):
powershell.exe -NoProfile -File .\load-csv-to-sql.ps1 `
    -ZipPath D:\deploy\data.zip -Table dbo.MyData -DryRun
```
