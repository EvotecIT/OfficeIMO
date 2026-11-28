# OfficeIMO.Notebook.Spreadsheet

Extract the saved data from a notebook output that contains
`application/vnd.dataresource+json`, such as a pandas table-schema display.
Select an output from the notebook, then export its inline rows to CSV or XLSX.

## Install

```sh
dotnet add package OfficeIMO.Notebook.Spreadsheet
```

```csharp
using OfficeIMO.Notebook;
using OfficeIMO.Notebook.Spreadsheet;

using var notebook = NotebookDocument.Load("analysis.ipynb");
var table = notebook.Cells.SelectMany(cell => cell.Outputs)
    .First(output => output.Data.ContainsKey("application/vnd.dataresource+json"));
using var destination = File.Create("measurements.xlsx");
var report = table.SaveAsExcel(destination);
```

Columns follow the saved schema order. Nulls and missing values become blank
cells. Excel stores numbers outside its supported precision as text. CSV
keeps the original spelling of numeric values and escapes formula-like text
by default.
Conversion requires inline scalar records and applies row, column and cell
limits. Presentation-only HTML tables use the document export routes.

