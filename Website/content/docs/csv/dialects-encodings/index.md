---
title: CSV Dialects and Encodings
description: Read and write semicolon, tab, and pipe-delimited files, detect delimiters, choose encodings and cultures, and apply strict or spreadsheet policy profiles with OfficeIMO.CSV and PSWriteOffice.
order: 53
packages:
  - OfficeIMO.CSV
cmdlets:
  - Import-OfficeCsv
limits: "Delimiter detection samples at most 64 Ki characters and 64 records; it falls back to the configured delimiter when the sample is inconclusive."
search_terms:
  - semicolon csv
  - tab separated
  - tsv
  - csv encoding
commands:
  - officeimo tabular convert
verified: 2026-11-14
---

"CSV" covers many dialects: comma or semicolon delimiters, quoted or unquoted fields, UTF-8 with or without a byte order mark, regional decimal separators. OfficeIMO.CSV reads and writes each of them explicitly, so the file you write back matches what the receiving system expects.

## Detect or set the delimiter

### C#

```csharp
using System.Text;
using OfficeIMO.CSV;

// Detect the delimiter from the first records
CsvDocument detected = CsvDocument.Load("export.csv", new CsvLoadOptions { DetectDelimiter = true });

// Or set it explicitly
CsvDocument semicolon = CsvDocument.Load("export-pl.csv", new CsvLoadOptions {
    Delimiter = ';',
    Encoding = Encoding.UTF8,
    HasHeaderRow = true
});
```

### PowerShell

```powershell
Import-OfficeCsv -Path .\export.csv -DetectDelimiter
Import-OfficeCsv -Path .\export-pl.csv -Delimiter ';'
```

### CLI

```sh
officeimo tabular convert export-pl.csv export.csv --delimiter ';' --output-delimiter ','
```

## Write for a regional spreadsheet

Use the culture that the reader expects for numbers and dates, and the matching delimiter.

### C#

```csharp
using System.Globalization;
using OfficeIMO.CSV;

var rows = new[] { new { Product = "Widget", Price = 9.99m } };

CsvDocument.FromObjects(rows, delimiter: ';', culture: new CultureInfo("de-DE"))
    .Save("prices-de.csv");
```

### PowerShell

```powershell
$data | Export-OfficeCsv -Path .\export.csv -UseCulture -Culture pl-PL
```

## Use a policy profile

Profiles set a consistent group of input and output rules in one call:

| Profile | Use it for |
|---|---|
| `Strict` | Machine-to-machine data. Malformed quotes, duplicate headers, and invalid UTF-8 fail. |
| `Compatibility` | Messy real-world files. Pads short rows, ignores extras, renames duplicate headers. |
| `Spreadsheet` | Files people open in Excel. Writes UTF-8 with BOM and CRLF, and escapes formula-like text. |
| `Tsv` | Tab-separated values without quoting, one record per line. |

### C#

```csharp
using OfficeIMO.CSV;

CsvLoadOptions input = CsvProfiles.CreateLoadOptions(CsvProfile.Strict);
input.Delimiter = ';';
CsvDocument document = CsvDocument.Load("input.csv", input);

document.Save("for-excel.csv", CsvProfiles.CreateSaveOptions(CsvProfile.Spreadsheet));
document.Save("copy.tsv", CsvProfiles.CreateSaveOptions(CsvProfile.Tsv));
```

## Control the output encoding

### C#

```csharp
using System.Text;
using OfficeIMO.CSV;

CsvDocument.Load("input.csv").Save("output.csv", new CsvSaveOptions {
    Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
    NewLine = "\n"
});
```

### PowerShell

```powershell
Import-OfficeCsv -Path .\input.csv | Export-OfficeCsv -Path .\output.csv -Encoding ([System.Text.UTF8Encoding]::new($false)) -NewLine "`n"
```
