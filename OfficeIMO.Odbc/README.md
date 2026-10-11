# OfficeIMO.Odbc

OfficeIMO.Odbc lets Windows x64 applications query an `.xlsx` workbook through ODBC. Each worksheet becomes a read-only table, with its first row supplying column names. The driver takes a snapshot when the connection opens; reconnect to see changes made to the workbook.

OfficeIMO.Excel reads the workbook. DbaClientX supplies an embedded SQLite engine for SQL queries against the snapshot. Users do not need Microsoft Office, the Microsoft Excel ODBC driver, a database server, or a SQLite database file.

## Build and register

Build on Windows x64 with the repository's .NET SDK and Visual Studio Build Tools' **Desktop development with C++** workload, which NativeAOT requires:

```powershell
dotnet publish OfficeIMO.Odbc/OfficeIMO.Odbc.csproj -c Release -r win-x64 -p:OfficeIMOOdbcNative=true -o artifacts/odbc
```

Deploy `OfficeIMO.Odbc.dll` and `e_sqlite3.dll` together in a trusted directory. The PDB and XML files are optional debugging and API documentation artifacts. The driver is a native shared library; applications do not need a .NET installation to load it.

Windows resolves the `DRIVER` connection attribute through its driver registry. Register the deployed DLL from an elevated **64-bit PowerShell** session, choosing a stable installation path:

```powershell
$driverPath = (Resolve-Path './artifacts/odbc/OfficeIMO.Odbc.dll').Path
$driverKey = 'HKLM:\SOFTWARE\ODBC\ODBCINST.INI\OfficeIMO Excel'
New-Item -Path $driverKey -ErrorAction Stop
New-ItemProperty -Path $driverKey -Name Driver -Value $driverPath -PropertyType String
New-ItemProperty -Path $driverKey -Name APILevel -Value '1' -PropertyType String
New-ItemProperty -Path $driverKey -Name DriverODBCVer -Value '03.00' -PropertyType String
New-ItemProperty -Path $driverKey -Name ConnectFunctions -Value 'NYN' -PropertyType String
New-ItemProperty -Path 'HKLM:\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers' -Name 'OfficeIMO Excel' -Value Installed -PropertyType String
```

The commands fail if that driver registration already exists. Inspect its current path before replacing it. To uninstall, close applications that use the driver, remove the `OfficeIMO Excel` driver key and its matching value under `ODBC Drivers`, then remove the deployed files. NativeAOT libraries stay loaded until their host process exits.

## Connect and query

Use a connection string with an absolute local workbook path:

```text
DRIVER={OfficeIMO Excel};DBQ={C:\Data\Customers.xlsx};ReadOnly=1;
```

Braces protect paths containing semicolons; double any `}` inside a braced value. Use `SQLDriverConnect` with `SQL_DRIVER_NOPROMPT`. The connection attributes are `DRIVER`, `DBQ`, and optional `ReadOnly=1`. Registered DSNs and connection dialogs are not supported.

For a worksheet named `Customers` with `Name`, `Country`, and `Amount` headers:

```sql
SELECT "Name", "Amount"
FROM "Customers"
WHERE "Country" = ?
ORDER BY "Amount" DESC
```

Queries use SQLite's SELECT syntax and positional ODBC `?` parameters. Quote worksheet and column identifiers with double quotes. Table names are the worksheet names, without an Excel `$` suffix.

An application using `System.Data.Odbc` can execute that query as follows:

```csharp
using System.Data.Odbc;

using var connection = new OdbcConnection(
    @"DRIVER={OfficeIMO Excel};DBQ={C:\Data\Customers.xlsx};ReadOnly=1;");
connection.Open();
using var command = connection.CreateCommand();
command.CommandText = "SELECT \"Name\", \"Amount\" FROM \"Customers\" WHERE \"Country\" = ?";
command.Parameters.Add("country", OdbcType.NVarChar, 32).Value = "Poland";
using var rows = command.ExecuteReader();
while (rows.Read())
    Console.WriteLine($"{rows.GetString(0)}: {rows.GetDouble(1)}");
```

`SQLTables`, `SQLColumns`, `SQLDescribeCol`, and `SQLGetTypeInfo` expose metadata for import tools. Unicode and ANSI entry points, input parameters, bound columns, chunked column reads, forward-only cursors, and cancellation are supported. ANSI strings use the Windows ANSI code page; use the Unicode APIs to preserve all characters.

## Data and limits

Column types are inferred across the complete worksheet. Numeric columns become doubles, Boolean columns become bits, dates become timestamps, and mixed columns become invariant text. Missing cells are SQL NULL; an explicit empty text cell is an empty string. Formula values use the workbook's saved results; the driver does not calculate formulas or run macros.

OfficeIMO.Excel supplies generated names for empty headers and disambiguates duplicate headers. The resulting column names must be unique ignoring case and at most 255 characters. Worksheet names must be unique ignoring case and must not start with `sqlite_`. The driver accepts `.xlsx` files on local absolute paths.

| Limit | Value |
| --- | --- |
| Source workbook | 64 MiB |
| Worksheets | 64 |
| Columns per worksheet or query result | 256 |
| Imported data rows across the workbook | 100,000 |
| Imported cells across the workbook | 1,000,000 |
| Imported text across the workbook | 32 Mi characters |
| SQL statement | 64 Ki characters; one SELECT |
| Parameters | 256 scalar input values |
| Result rows | 100,000, or a lower `SQL_ATTR_MAX_ROWS` |
| Snapshot import deadline | 60 seconds |
| Query deadline | 30 seconds by default; configurable up to 300 seconds |
| Open driver connections | 32 per process |
| Driver handles | 4,096 per process |

Driver calls are serialized within a process; `SQLCancel` can interrupt a running query from another thread. Each statement has its own query reader. The driver buffers workbook data for import, so it is intended for bounded spreadsheet imports rather than large database workloads.

The driver rejects writes, transactions, SQL batches, CTEs, PRAGMA/ATTACH commands, stored procedures, asynchronous ODBC execution, parameter arrays, explicit descriptors, scrollable cursors, and row arrays larger than one. It does not translate another database's SQL dialect or schema conventions. Compatibility with a specific business application depends on the ODBC operations and SQL that application uses. Entity Framework needs an appropriate provider; an ODBC driver alone does not supply one.

## Verify

The optional driver and its verification projects stay outside the main solution so ordinary OfficeIMO consumers do not acquire SQLite. Run managed contracts, publish the native driver, build the independent .NET client, then exercise both clients through Windows' driver manager:

```powershell
dotnet test OfficeIMO.Odbc.Tests/OfficeIMO.Odbc.Tests.csproj -c Release
dotnet publish OfficeIMO.Odbc/OfficeIMO.Odbc.csproj -c Release -r win-x64 -p:OfficeIMOOdbcNative=true -o artifacts/odbc
dotnet build OfficeIMO.Odbc.Verification/OfficeIMO.Odbc.Verification.csproj -c Release -o artifacts/odbc-client
python OfficeIMO.Odbc.Verification/native_smoke.py --driver artifacts/odbc/OfficeIMO.Odbc.dll --ado-consumer artifacts/odbc-client/OfficeIMO.Odbc.Verification.dll
```

The verification creates its own workbook and a private registry hive, redirects registry lookup only inside its child test processes, and removes both afterward. It does not require administrator rights or modify installed ODBC drivers. It checks Unicode, numeric and null values, parameter binding, column binding, metadata, truncation, cancellation, reconnection, and that the source workbook remains unchanged.
