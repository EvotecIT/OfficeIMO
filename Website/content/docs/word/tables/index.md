---
title: Add Tables to Word Documents
description: "Create Word tables from .NET objects or cell by cell, apply table styles, merge cells, set widths and borders, and repeat header rows with OfficeIMO.Word."
order: 12
packages:
  - OfficeIMO.Word
cmdlets:
  - Add-OfficeWordTable
  - Add-OfficeWordTableCondition
  - Get-OfficeWordTable
search_terms:
  - word table
  - docx table
  - table from objects
  - merge cells
  - table style
verified: 2026-11-14
---

`WordTable` is a native Word table. Build one from a list of objects when the data already exists in your application, or by row and column count when you need full control over each cell. The `WordTableStyle` enum exposes Word's built-in table styles (`TableGrid`, `GridTable4Accent1`, `ListTable3Accent2`, and the rest of the gallery).

## Add a table from a list of objects

Each public property becomes a column, and the property names become the header row. `WordTableOptions` adds a style, a caption, and column selection on top of the shared `OfficeTableOptions`.

### C#

```csharp
using OfficeIMO.Word;

Employee[] employees = [
    new("Alice", "Engineering", 95_000),
    new("Bob", "Design", 85_000),
    new("Carol", "Operations", 105_000)
];

using WordDocument document = WordDocument.Create("employees.docx");
document.AddParagraph("Employees").Style = WordParagraphStyles.Heading1;
document.AddTable(employees, new WordTableOptions {
    Style = WordTableStyle.GridTable4Accent1,
    Caption = "Table 1. Employees by department"
});
document.Save();

record Employee(string Name, string Department, int Salary);
```

### PowerShell

`Add-OfficeWordTableCondition` highlights rows that match a filter.

```powershell
$rows = Get-Service | Select-Object -First 10 Name, Status, StartType
New-OfficeWord -Path .\services.docx {
    Add-OfficeWordHeading 'Services'
    Add-OfficeWordTable -InputObject $rows -Style GridTable4Accent1 {
        Add-OfficeWordTableCondition -FilterScript { $_.Status -eq 'Stopped' } -BackgroundColor '#FDE2E1'
    }
}
```

## Build a table cell by cell

Every cell starts with one empty paragraph, so `Paragraphs[0]` is always there to write into. Add more paragraphs to a cell with `AddParagraph`.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("schedule.docx");

WordTable table = document.AddTable(3, 3, WordTableStyle.GridTable1Light);
string[,] values = {
    { "Name", "Department", "Start date" },
    { "Alice", "Engineering", "2026-01-15" },
    { "Bob", "Design", "2026-03-01" }
};
for (int r = 0; r < 3; r++) {
    for (int c = 0; c < 3; c++) {
        table.Rows[r].Cells[c].Paragraphs[0].Text = values[r, c];
    }
}
foreach (WordTableCell header in table.Rows[0].Cells) header.Paragraphs[0].Bold = true;

WordTableRow row = table.AddRow(3);
row.Cells[0].Paragraphs[0].Text = "Carol";
row.Cells[1].Paragraphs[0].Text = "Operations";
row.Cells[2].Paragraphs[0].Text = "2026-06-10";

document.Save();
```

## Merge cells

`MergeCells` takes the top-left cell and how many rows and columns to span.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("merged.docx");
WordTable table = document.AddTable(4, 4, WordTableStyle.TableGrid);

table.Rows[0].Cells[0].Paragraphs[0].Text = "2026 sales by quarter";
table.MergeCells(0, 0, rowSpan: 1, colSpan: 4);

table.Rows[1].Cells[0].Paragraphs[0].Text = "Hardware";
table.MergeCells(1, 0, rowSpan: 3, colSpan: 1);

document.Save();
```

## Repeat the header row and control page breaks

For long tables, repeat the first row at the top of every page and keep rows from splitting across a page break.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("long-table.docx");
WordTable table = document.AddTable(Enumerable.Range(1, 200)
    .Select(i => new[] { $"Item {i}", (i * 3).ToString() })
    .Prepend(new[] { "Item", "Quantity" }));

table.RepeatAsHeaderRowAtTheTopOfEachPage = true;
table.AllowRowToBreakAcrossPages = false;
table.Style = WordTableStyle.ListTable3Accent2;

document.Save();
```

## Set table width, layout, and borders

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using WordDocument document = WordDocument.Create("borders.docx");
WordTable table = document.AddTable(2, 3, WordTableStyle.TableGrid);
table.SetWidthPercentage(100);
table.LayoutMode = WordTableLayoutMode.Fixed;

WordTableCellBorder borders = table.Rows[0].Cells[0].Borders;
borders.BottomStyle = WordBorderStyle.Double;
borders.BottomSize = OfficeLength.Points(1.5);
borders.BottomColor = OfficeColor.Parse("1F4E79");

document.Save();
```

## Read tables from an existing document

`document.Tables` lists every top-level table. Each cell's text is available through its paragraphs.

### C#

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("employees.docx");
foreach (WordTable table in document.Tables) {
    foreach (WordTableRow row in table.Rows) {
        Console.WriteLine(string.Join(" | ", row.Cells.Select(cell =>
            string.Join(" ", cell.Paragraphs.Select(p => p.Text)))));
    }
}
```

### PowerShell

```powershell
Get-OfficeWordTable -Path .\employees.docx
```
