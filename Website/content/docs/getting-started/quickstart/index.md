---
title: Quick Start
description: Create your first Word document, Excel workbook, and PowerPoint presentation with OfficeIMO in minutes. Includes examples and API links.
order: 2
---

This guide walks you through creating simple documents with OfficeIMO. All examples assume you have already [installed](/docs/getting-started/installation) the relevant NuGet packages.

## Create a Word Document (C#)

```csharp
using OfficeIMO.Word;

// Create a new document saved to disk
using var document = WordDocument.Create("HelloWorld.docx");

// Add a paragraph with text
document.AddParagraph("Hello, World!").Bold = true;

// Add more content
var paragraph = document.AddParagraph("This document was created with OfficeIMO.");
paragraph.FontFamily = "Calibri";
paragraph.FontSize = 14;
paragraph.Color = OfficeIMO.OfficeColor.DarkBlue;

// Add a table
var table = document.AddTable(3, 3, WordTableStyle.TableGrid);
table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
table.Rows[0].Cells[1].Paragraphs[0].Text = "Role";
table.Rows[0].Cells[2].Paragraphs[0].Text = "Location";
table.Rows[1].Cells[0].Paragraphs[0].Text = "Alice";
table.Rows[1].Cells[1].Paragraphs[0].Text = "Developer";
table.Rows[1].Cells[2].Paragraphs[0].Text = "New York";
table.Rows[2].Cells[0].Paragraphs[0].Text = "Bob";
table.Rows[2].Cells[1].Paragraphs[0].Text = "Designer";
table.Rows[2].Cells[2].Paragraphs[0].Text = "London";

document.Save();
```

## Add document properties and styled text

```csharp
using OfficeIMO.Word;

using var document = WordDocument.Create("report.docx");
document.BuiltinDocumentProperties.Title = "My Report";
document.BuiltinDocumentProperties.Creator = "OfficeIMO";
var introduction = document.AddParagraph("Introduction");
introduction.Style = WordParagraphStyles.Heading1;
introduction.Bold = true;
document.AddParagraph("This report was generated programmatically.");
document.Save();
```

## Create an Excel Workbook (C#)

```csharp
using OfficeIMO.Excel;

// Create a workbook with one sheet
using var workbook = ExcelDocument.Create("Report.xlsx", "Sales");

// Access the first sheet
var sheet = workbook.Sheets[0];

// Set header row
sheet.CellAt("A1").SetValue("Product");
sheet.CellAt("B1").SetValue("Q1 Revenue");
sheet.CellAt("C1").SetValue("Q2 Revenue");

// Set data rows
sheet.CellAt("A2").SetValue("Widget A");
sheet.CellAt("B2").SetValue(15000);
sheet.CellAt("C2").SetValue(18000);

sheet.CellAt("A3").SetValue("Widget B");
sheet.CellAt("B3").SetValue(22000);
sheet.CellAt("C3").SetValue(25000);

// Add a SUM formula
sheet.CellAt("B4").SetFormula("SUM(B2:B3)");
sheet.CellAt("C4").SetFormula("SUM(C2:C3)");

workbook.Save();
```

## Create a Markdown Document (C#)

```csharp
using OfficeIMO.Markdown;

var doc = MarkdownDocument.Create()
    .H1("Project Status Report")
    .P("Generated automatically by OfficeIMO.Markdown.")
    .H2("Summary")
    .Table(t => t
        .Headers("Task", "Status", "Owner")
        .Row("Backend API", "Complete", "Alice")
        .Row("Frontend UI", "In Progress", "Bob")
        .Row("Documentation", "Planned", "Carol")
    )
    .H2("Notes")
    .Ul(ul => ul
        .Item("Sprint ends Friday")
        .Item("Demo scheduled for Monday")
    );

var markdown = doc.ToMarkdown();
File.WriteAllText("status.md", markdown);
```

## Create a CSV Document (C#)

```csharp
using OfficeIMO.CSV;

var csv = new CsvDocument()
    .WithDelimiter(',')
    .WithHeader("Name", "Age", "City")
    .AddRow("Alice", "30", "New York")
    .AddRow("Bob", "25", "London")
    .AddRow("Carol", "35", "Tokyo");

csv.Save("people.csv");
```

Or generate a CSV from objects:

```csharp
var people = new[] {
    new { Name = "Alice", Age = 30, City = "New York" },
    new { Name = "Bob", Age = 25, City = "London" },
};

var csv = CsvDocument.FromObjects(people);
csv.Save("people.csv");
```

## PowerShell Example

```powershell
Import-Module PSWriteOffice

$doc = New-OfficeWord -Path "Report.docx" -PassThru

$doc | Add-OfficeWordSection {
    Add-OfficeWordParagraph -Text "Hello from PowerShell!" -Style Heading1
}
$doc | Add-OfficeWordParagraph {
    Add-OfficeWordText -Text "Created with PSWriteOffice." -Bold
}

$doc | Save-OfficeWord
Close-OfficeWord -Document $doc

$excel = New-OfficeExcel -Path "Data.xlsx" -PassThru

$excel | Add-OfficeExcelSheet -Name "Summary" {
    Set-OfficeExcelCell -Address "A1" -Value "Metric"
    Set-OfficeExcelCell -Address "B1" -Value "Value"
}

$excel | Save-OfficeExcel
Close-OfficeExcel -Document $excel
```

## Next Steps

- Dive deeper into [Word documents](/docs/word/)
- Explore [Excel workbooks](/docs/excel/)
- Learn about [Markdown generation](/docs/markdown/)
- Check [platform support](/docs/getting-started/platform-support) for your target environment
