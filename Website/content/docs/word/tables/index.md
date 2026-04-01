---
title: Tables
description: Creating and styling tables in OfficeIMO.Word with rows, cells, merging, borders, and the built-in styles exposed by the WordTableStyle enum.
order: 12
---

# Tables

The `WordTable` class provides the table-focused API for creating and manipulating tables in Word documents. The current repo surface includes the built-in styles exposed by `WordTableStyle`, plus cell merging, custom borders, header row repetition, and automatic table generation from object collections.

## Creating a Table

```csharp
using OfficeIMO.Word;

using var document = WordDocument.Create("tables.docx");

// Create a 4-row, 3-column table with the TableGrid style
var table = document.AddTable(4, 3, WordTableStyle.TableGrid);

// Populate the header row
table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
table.Rows[0].Cells[1].Paragraphs[0].Text = "Department";
table.Rows[0].Cells[2].Paragraphs[0].Text = "Start Date";

// Populate data rows
table.Rows[1].Cells[0].Paragraphs[0].Text = "Alice";
table.Rows[1].Cells[1].Paragraphs[0].Text = "Engineering";
table.Rows[1].Cells[2].Paragraphs[0].Text = "2024-01-15";

table.Rows[2].Cells[0].Paragraphs[0].Text = "Bob";
table.Rows[2].Cells[1].Paragraphs[0].Text = "Design";
table.Rows[2].Cells[2].Paragraphs[0].Text = "2024-03-01";

table.Rows[3].Cells[0].Paragraphs[0].Text = "Carol";
table.Rows[3].Cells[1].Paragraphs[0].Text = "Marketing";
table.Rows[3].Cells[2].Paragraphs[0].Text = "2024-06-10";

document.Save();
```

## Table Styles

OfficeIMO includes over 105 built-in table styles accessible through the `WordTableStyle` enum. Some commonly used styles:

| Style | Description |
|-------|-------------|
| `WordTableStyle.TableGrid` | Simple grid with all borders |
| `WordTableStyle.PlainTable1` | Minimal borders, clean look |
| `WordTableStyle.GridTable1Light` | Light grid with subtle shading |
| `WordTableStyle.GridTable4Accent1` | Colored banded rows with accent color 1 |
| `WordTableStyle.ListTable3Accent2` | List-style with accent color 2 |
| `WordTableStyle.TableGridLight` | Very light grid borders |

Apply a style when creating:

```csharp
var table = document.AddTable(3, 3, WordTableStyle.GridTable4Accent1);
```

Or change the style after creation:

```csharp
table.Style = WordTableStyle.ListTable3Accent2;
```

## Adding Rows and Cells

```csharp
// Add a new row to an existing table
var row = table.AddRow();
row.Cells[0].Paragraphs[0].Text = "New entry";
row.Cells[1].Paragraphs[0].Text = "Operations";
row.Cells[2].Paragraphs[0].Text = "2025-01-01";
```

## Cell Content

Each cell contains a list of paragraphs. You can add multiple paragraphs to a cell, format text within cells, and even nest tables:

```csharp
var cell = table.Rows[1].Cells[0];

// The first paragraph is created automatically
cell.Paragraphs[0].Text = "First line";
cell.Paragraphs[0].Bold = true;

// Add more paragraphs to the same cell
var secondParagraph = cell.AddParagraph();
secondParagraph.Text = "Second line";
secondParagraph.Italic = true;
```

## Header Row Repetition

For long tables that span multiple pages, repeat the header row at the top of each page:

```csharp
table.RepeatAsHeaderRowAtTheTopOfEachPage = true;
```

## Row Break Across Pages

Control whether rows can break across page boundaries:

```csharp
// Allow all rows to break across pages
table.AllowRowToBreakAcrossPages = true;

// Or control individual rows
table.Rows[2].AllowRowToBreakAcrossPages = false;
```

## Cell Merging

### Horizontal Merge (across columns)

```csharp
// Merge cells 0-2 in the first row (spans 3 columns)
table.Rows[0].Cells[0].HorizontalMerge = MergedCellValues.Restart;
table.Rows[0].Cells[1].HorizontalMerge = MergedCellValues.Continue;
table.Rows[0].Cells[2].HorizontalMerge = MergedCellValues.Continue;
```

### Vertical Merge (across rows)

```csharp
// Merge first column across rows 1-3
table.Rows[1].Cells[0].VerticalMerge = MergedCellValues.Restart;
table.Rows[2].Cells[0].VerticalMerge = MergedCellValues.Continue;
table.Rows[3].Cells[0].VerticalMerge = MergedCellValues.Continue;
```

## Table Borders

```csharp
// Access and customize borders via WordTableBorders
var borders = table.Borders;
borders.TopBorder = new WordBorder {
    Value = BorderValues.Single,
    Size = 12,
    Color = "FF0000"
};
borders.BottomBorder = new WordBorder {
    Value = BorderValues.Double,
    Size = 6,
    Color = "0000FF"
};
```

## Table Layout

Control whether the table uses a fixed or auto-fit layout:

```csharp
table.LayoutType = WordTableLayoutType.Fixed;
// or
table.LayoutType = WordTableLayoutType.Autofit;
```

## Table Width

```csharp
table.Width = 5000;     // in fiftieths of a percent or twips
table.WidthType = TableWidthUnitValues.Pct;
```

## Generate Tables from Objects

Create a table automatically from a collection of objects:

```csharp
var employees = new[] {
    new { Name = "Alice", Role = "Developer", Salary = 95000 },
    new { Name = "Bob", Role = "Designer", Salary = 85000 },
    new { Name = "Carol", Role = "Manager", Salary = 105000 },
};

var table = document.AddTableFromObjects(
    employees,
    WordTableStyle.GridTable4Accent1,
    includeHeader: true
);
```

## Accessing All Table Paragraphs

Iterate over every paragraph in a table (useful for search/replace):

```csharp
foreach (var paragraph in table.Paragraphs) {
    if (paragraph.Text.Contains("placeholder")) {
        paragraph.Text = paragraph.Text.Replace("placeholder", "actual value");
    }
}
```
