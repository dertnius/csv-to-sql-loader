<#
    Pester unit tests for deploy/load-csv-to-sql.ps1.

    Dot-sources the script (its main-guard prevents the load from running; ZipPath/Table are
    deliberately non-mandatory so dot-sourcing never prompts) and tests the pure helpers in
    isolation -- no SQL Server, no SqlClient provider, no files touched on disk beyond a temp
    CSV for the Import-Csv-shaped round-trip.

    Run:  pwsh -NoProfile -Command "Invoke-Pester -Path deploy/tests -CI"
      or: powershell.exe -NoProfile -Command "Invoke-Pester -Path deploy\tests -CI"
#>

BeforeAll {
    # [IO.Path]::Combine handles 3+ segments on BOTH pwsh 7 and Windows PowerShell 5.1.
    # (5.1's Join-Path takes only -Path/-ChildPath -- a 3rd positional arg is a bind error.)
    $script:ScriptPath = (Resolve-Path ([IO.Path]::Combine($PSScriptRoot, '..', 'load-csv-to-sql.ps1'))).Path
    . $script:ScriptPath          # main-guard => functions only, no execution
}

Describe 'Test-SqlIdentifier' {
    It 'accepts schema.table and a bare table' {
        Test-SqlIdentifier -Name 'dbo.MyData' | Should -BeTrue
        Test-SqlIdentifier -Name 'MyData'     | Should -BeTrue
        Test-SqlIdentifier -Name '_staging'   | Should -BeTrue
    }
    It 'rejects spaces, brackets, quotes and injection attempts' {
        Test-SqlIdentifier -Name 'my table'            | Should -BeFalse
        Test-SqlIdentifier -Name '[dbo].[MyData]'      | Should -BeFalse
        Test-SqlIdentifier -Name 'x];DROP TABLE y--'   | Should -BeFalse
        Test-SqlIdentifier -Name "a';b"                | Should -BeFalse
    }
    It 'rejects a three-part name and a leading digit' {
        Test-SqlIdentifier -Name 'db.schema.table' | Should -BeFalse
        Test-SqlIdentifier -Name '1table'          | Should -BeFalse
    }
}

Describe 'Split-SqlTable / Get-QuotedTableName' {
    It 'splits schema.table' {
        $p = Split-SqlTable -Name 'dbo.MyData'
        $p.Schema | Should -Be 'dbo'
        $p.Table  | Should -Be 'MyData'
    }
    It 'defaults the schema to dbo for a bare table' {
        (Split-SqlTable -Name 'MyData').Schema | Should -Be 'dbo'
        (Split-SqlTable -Name 'MyData').Table  | Should -Be 'MyData'
    }
    It 'quotes both parts with brackets' {
        Get-QuotedTableName -Parts (Split-SqlTable -Name 'app.Data') | Should -Be '[app].[Data]'
    }
}

Describe 'Format-ConnValue' {
    It 'leaves a plain value unquoted' {
        Format-ConnValue -Value 'localhost'     | Should -Be 'localhost'
        Format-ConnValue -Value 'HOST\SQL2019'  | Should -Be 'HOST\SQL2019'
    }
    It 'double-quotes a value with a semicolon or equals' {
        Format-ConnValue -Value 'a;b' | Should -Be '"a;b"'
        Format-ConnValue -Value 'a=b' | Should -Be '"a=b"'
    }
    It 'single-quotes a value that itself contains a double quote' {
        Format-ConnValue -Value 'a"b;c' | Should -Be "'a`"b;c'"
    }
}

Describe 'New-SqlConnectionString' {
    It 'uses Windows integrated auth and carries no credentials' {
        $cs = New-SqlConnectionString -Server 'localhost' -Database 'database'
        $cs | Should -Match 'Integrated Security=SSPI'
        $cs | Should -Match 'Data Source=localhost'
        $cs | Should -Match 'Initial Catalog=database'
        $cs | Should -Not -Match 'Password'
        $cs | Should -Not -Match 'User ID'
    }
    It 'preserves a named instance' {
        (New-SqlConnectionString -Server 'HOST\SQL2019' -Database 'database') | Should -Match 'Data Source=HOST\\SQL2019'
    }
}

Describe 'Get-CsvColumns' {
    It 'returns header names in order from the first record' {
        $recs = @([pscustomobject]@{ A = '1'; B = '2'; C = '3' })
        (Get-CsvColumns -Records $recs -HeaderLine 'A,B,C' -Delimiter ',') -join ',' | Should -Be 'A,B,C'
    }
    It 'falls back to the header line (BOM-stripped) for a header-only CSV' {
        $header = "$([char]0xFEFF)Col1,Col2,Col3"
        (Get-CsvColumns -Records @() -HeaderLine $header -Delimiter ',') -join ',' | Should -Be 'Col1,Col2,Col3'
    }
    It 'honours a non-comma delimiter in the fallback' {
        (Get-CsvColumns -Records @() -HeaderLine 'A;B;C' -Delimiter ';') -join ',' | Should -Be 'A,B,C'
    }
}

Describe 'Get-MissingColumns' {
    It 'reports CSV columns absent from the table' {
        (Get-MissingColumns -CsvColumns @('A', 'b', 'Z') -TableColumns @('a', 'B', 'c')) -join ',' | Should -Be 'Z'
    }
    It 'is case-insensitive and returns nothing when all map' {
        @(Get-MissingColumns -CsvColumns @('Name', 'AGE') -TableColumns @('name', 'age', 'extra')).Count | Should -Be 0
    }
}

Describe 'Resolve-CsvFile' {
    It 'returns the single CSV in the archive' {
        Resolve-CsvFile -Files @('C:\x\data.csv', 'C:\x\readme.txt') | Should -Be 'C:\x\data.csv'
    }
    It 'throws when there is no CSV' {
        { Resolve-CsvFile -Files @('C:\x\a.txt') } | Should -Throw -ExpectedMessage '*No .csv*'
    }
    It 'throws when there are several and none is named' {
        { Resolve-CsvFile -Files @('C:\x\a.csv', 'C:\x\b.csv') } | Should -Throw -ExpectedMessage '*Multiple*'
    }
    It 'selects a named CSV when several are present' {
        Resolve-CsvFile -Files @('C:\x\a.csv', 'C:\x\b.csv') -CsvName 'b.csv' | Should -Be 'C:\x\b.csv'
    }
    It 'throws when the named CSV is absent' {
        { Resolve-CsvFile -Files @('C:\x\a.csv') -CsvName 'missing.csv' } | Should -Throw -ExpectedMessage "*not found*"
    }
}

Describe 'ConvertTo-DataTable' {
    It 'builds a DataTable with one row per record and the given columns' {
        $recs = @(
            [pscustomobject]@{ A = '1'; B = 'x' },
            [pscustomobject]@{ A = '2'; B = 'y' }
        )
        $dt = ConvertTo-DataTable -Records $recs -Columns @('A', 'B')
        # Note: don't pipe $dt into Should -- a DataTable enumerates to its DataRows in a pipeline.
        $dt.GetType().Name | Should -Be 'DataTable'
        $dt.Rows.Count | Should -Be 2
        $dt.Columns.Count | Should -Be 2
        $dt.Rows[1]['A'] | Should -Be '2'
    }
    It 'maps empty strings to DBNull so nullable typed columns get NULL' {
        $recs = @([pscustomobject]@{ A = ''; B = 'kept' })
        $dt = ConvertTo-DataTable -Records $recs -Columns @('A', 'B')
        $dt.Rows[0]['A'] | Should -BeOfType [System.DBNull]
        $dt.Rows[0]['B'] | Should -Be 'kept'
    }
    It 'returns the table object itself, not its enumerated rows' {
        # The `,$dt` return in ConvertTo-DataTable matters: Windows PowerShell 5.1 enumerates a
        # DataTable to its DataRows on the pipeline (PS7 does not), so a plain `return $dt` would
        # hand the caller a bare DataRow. A reachable .Rows proves the object is still the table.
        $dt = ConvertTo-DataTable -Records @([pscustomobject]@{ A = '1' }) -Columns @('A')
        $dt.GetType().Name | Should -Be 'DataTable'
        $dt.Rows.Count | Should -Be 1
    }
    It 'round-trips real CSV headers from Import-Csv' {
        $tmp = Join-Path ([IO.Path]::GetTempPath()) ("csvload-" + [Guid]::NewGuid().ToString('N') + ".csv")
        try {
            "Id,Name,Notes`n1,Alice,`n2,Bob,hi" | Set-Content -LiteralPath $tmp -Encoding utf8
            $recs = @(Import-Csv -LiteralPath $tmp)
            $cols = Get-CsvColumns -Records $recs -Delimiter ','
            $cols -join ',' | Should -Be 'Id,Name,Notes'
            $dt = ConvertTo-DataTable -Records $recs -Columns $cols
            $dt.Rows.Count | Should -Be 2
            $dt.Rows[0]['Notes'] | Should -BeOfType [System.DBNull]   # empty trailing cell -> NULL
            $dt.Rows[1]['Name']  | Should -Be 'Bob'
        }
        finally { if (Test-Path $tmp) { Remove-Item $tmp -Force } }
    }
}
