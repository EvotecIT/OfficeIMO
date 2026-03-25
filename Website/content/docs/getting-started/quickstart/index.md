---
title: Quick Start
description: Create your first Word document, Excel workbook, and PowerPoint presentation with OfficeIMO in minutes.
order: 2
---

# Quick Start

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
paragraph.Color = SixLabors.ImageSharp.Color.DarkBlue;

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

## Create a Word Document with the Fluent API

OfficeIMO.Word also provides a fluent builder API via `WordFluentDocument`:

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

using var document = WordFluentDocument.Create("Fluent.docx")
    .Info(i => i.Title("My Report").Author("OfficeIMO"))
    .Section(s => s
        .Paragraph(p => p.Text("Introduction").Bold().FontSize(16))
        .Paragraph(p => p.Text("This report was generated programmatically."))
    )
    .Build();

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
sheet.Cells["A1"].Value = "Product";
sheet.Cells["B1"].Value = "Q1 Revenue";
sheet.Cells["C1"].Value = "Q2 Revenue";

// Set data rows
sheet.Cells["A2"].Value = "Widget A";
sheet.Cells["B2"].Value = 15000;
sheet.Cells["C2"].Value = 18000;

sheet.Cells["A3"].Value = "Widget B";
sheet.Cells["B3"].Value = 22000;
sheet.Cells["C3"].Value = 25000;

// Add a SUM formula
sheet.Cells["B4"].Value = "=SUM(B2:B3)";
sheet.Cells["C4"].Value = "=SUM(C2:C3)";

workbook.Save();
```

## Create a Markdown Document (C#)

```csharp
using OfficeIMO.Markdown;

var doc = MarkdownDoc.Create()
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

# Create a Word document
$doc = New-OfficeWord -Path "Report.docx" -PassThru

$doc | Add-OfficeWordSection {
    Add-OfficeWordParagraph -Text "Hello from PowerShell!" -Style Heading1
}
$doc | Add-OfficeWordParagraph {
    Add-OfficeWordText -Text "Created with PSWriteOffice." -Bold
}

$doc | Save-OfficeWord
Close-OfficeWord -Document $doc

# Create an Excel workbook
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
