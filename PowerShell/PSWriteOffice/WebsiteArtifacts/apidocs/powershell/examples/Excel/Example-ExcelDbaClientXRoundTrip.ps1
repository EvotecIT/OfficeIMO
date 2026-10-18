param(
    [string] $Server = $(if ($env:DBACLIENTX_SQLSERVER) { $env:DBACLIENTX_SQLSERVER } else { 'localhost' }),
    [string] $Database = $(if ($env:DBACLIENTX_SQLDATABASE) { $env:DBACLIENTX_SQLDATABASE } else { 'tempdb' }),
    [string] $SourceTable = 'dbo.PSWriteOfficeExcelSource',
    [string] $DestinationTable = 'dbo.PSWriteOfficeExcelRoundTrip',
    [string] $WorksheetName = 'Rows',
    [string] $Path,
    [int] $RowCount = 100,
    [switch] $KeepArtifacts
)

Import-Module PSWriteOffice -ErrorAction Stop
Import-Module DbaClientX -ErrorAction Stop

if ($RowCount -lt 1) {
    throw 'RowCount must be greater than zero.'
}

if (-not $Path) {
    $outputDirectory = Join-Path $PSScriptRoot '..\Documents'
    if (-not (Test-Path -LiteralPath $outputDirectory)) {
        $null = New-Item -Path $outputDirectory -ItemType Directory -Force
    }

    $Path = Join-Path $outputDirectory 'Excel-DbaClientXRoundTrip.xlsx'
}

$connectionString = "Server=$Server;Database=$Database;Encrypt=True;TrustServerCertificate=True;Integrated Security=True"

Invoke-DbaXNonQuery -Server $Server -Database $Database -TrustServerCertificate -Query @"
IF OBJECT_ID(N'$SourceTable', N'U') IS NOT NULL DROP TABLE $SourceTable;
IF OBJECT_ID(N'$DestinationTable', N'U') IS NOT NULL DROP TABLE $DestinationTable;

CREATE TABLE $SourceTable
(
    Id int NOT NULL,
    Department nvarchar(100) NOT NULL,
    Amount decimal(18,2) NOT NULL,
    CreatedUtc datetime2 NOT NULL
);

WITH numbers AS
(
    SELECT TOP ($RowCount)
        ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS Id
    FROM sys.all_objects AS a
    CROSS JOIN sys.all_objects AS b
)
INSERT INTO $SourceTable (Id, Department, Amount, CreatedUtc)
SELECT
    Id,
    CONCAT(N'Department ', ((Id - 1) % 5) + 1),
    CONVERT(decimal(18,2), Id * 10.25),
    SYSUTCDATETIME()
FROM numbers;
"@ -ErrorAction Stop | Out-Null

try {
    $rows = @(
        Invoke-DbaXQuery `
            -Server $Server `
            -Database $Database `
            -TrustServerCertificate `
            -Query "SELECT Id, Department, Amount, CreatedUtc FROM $SourceTable ORDER BY Id;" `
            -ReturnType PSObject `
            -ErrorAction Stop
    )

    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force
    }

    $rows |
        Export-OfficeExcel `
            -Path $Path `
            -WorksheetName $WorksheetName `
            -TableName 'DatabaseRows' `
            -AutoFit `
            -FreezeTopRow `
            -BoldTopRow `
            -ErrorAction Stop

    $table = Import-OfficeExcel `
        -Path $Path `
        -WorksheetName $WorksheetName `
        -AsDataTable `
        -ErrorAction Stop

    $writeResult = $table |
        Write-DbaXTableData `
            -Provider SqlServer `
            -ConnectionString $connectionString `
            -DestinationTable $DestinationTable `
            -AutoCreateTable `
            -TableLock `
            -BatchSize 5000 `
            -PassThru `
            -ErrorAction Stop

    $verification = Invoke-DbaXQuery `
        -Server $Server `
        -Database $Database `
        -TrustServerCertificate `
        -Query "SELECT COUNT(*) AS SourceRows FROM $SourceTable; SELECT COUNT(*) AS DestinationRows FROM $DestinationTable;" `
        -ReturnType DataSet `
        -ErrorAction Stop

    $sourceRows = [int] $verification.Tables[0].Rows[0]['SourceRows']
    $destinationRows = [int] $verification.Tables[1].Rows[0]['DestinationRows']

    $hasMismatch = $rows.Count -ne $RowCount -or
        $table.Rows.Count -ne $RowCount -or
        $writeResult.Rows -ne $RowCount -or
        $sourceRows -ne $RowCount -or
        $destinationRows -ne $RowCount

    if ($hasMismatch) {
        throw "Round-trip row count mismatch. Exported=$($rows.Count), Imported=$($table.Rows.Count), Written=$($writeResult.Rows), Source=$sourceRows, Destination=$destinationRows, Expected=$RowCount."
    }

    [pscustomobject]@{
        Server = $Server
        Database = $Database
        Path = $Path
        WorksheetName = $WorksheetName
        SourceTable = $SourceTable
        DestinationTable = $DestinationTable
        ExportedRows = $rows.Count
        ImportedRows = $table.Rows.Count
        WrittenRows = $writeResult.Rows
        SourceRows = $sourceRows
        DestinationRows = $destinationRows
    }
}
finally {
    if (-not $KeepArtifacts) {
        Invoke-DbaXNonQuery -Server $Server -Database $Database -TrustServerCertificate -Query @"
IF OBJECT_ID(N'$DestinationTable', N'U') IS NOT NULL DROP TABLE $DestinationTable;
IF OBJECT_ID(N'$SourceTable', N'U') IS NOT NULL DROP TABLE $SourceTable;
"@ -ErrorAction SilentlyContinue | Out-Null

        if (Test-Path -LiteralPath $Path) {
            Remove-Item -LiteralPath $Path -Force
        }
    }
}
