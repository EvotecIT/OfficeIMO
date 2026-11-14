---
title: Build Markdown Documents in Code
description: Write Markdown files from code with headings, formatted paragraphs, lists, task lists, tables, code blocks, callouts, images, details and front matter.
order: 41
search_terms:
  - markdown builder
  - generate markdown
  - markdown table from objects
  - write md file
  - markdown report
packages:
  - OfficeIMO.Markdown
cmdlets:
  - New-OfficeMarkdown
  - Add-OfficeMarkdownHeading
  - Add-OfficeMarkdownParagraph
  - Add-OfficeMarkdownList
  - Add-OfficeMarkdownTaskList
  - Add-OfficeMarkdownTable
  - Add-OfficeMarkdownCode
  - Add-OfficeMarkdownCallout
  - Add-OfficeMarkdownFrontMatter
  - Add-OfficeMarkdownTableOfContents
limits: Builder lists are flat; nested lists come from parsing Markdown text or adding ListItem blocks to the model directly.
verified: 2026-11-14
---

`MarkdownDocument` has two ways to add content. The chaining shortcuts (`H1`–`H6`, `P`, `Ul`, `Ol`, `Dl`, `Table`, `Code`, `Quote`, `Callout`, `Image`, `Details`) return the document so you can keep chaining. The `Add` methods (`AddHeading`, `AddParagraph`, `AddTable`, `AddList`) return the block they created so you can keep editing it. `ToMarkdown()` writes GitHub-flavored Markdown; `Save(path)` writes it to a file.

In PowerShell, `New-OfficeMarkdown -Path ... { }` runs the content block and saves the file when the block completes.

## Add headings and formatted paragraphs

### C#

```csharp
using OfficeIMO.Markdown;

MarkdownDocument document = MarkdownDocument.Create()
    .H1("Release notes")
    .H2("Highlights")
    .P("Plain paragraph text.")
    .P(p => p
        .Text("This is ")
        .Bold("important")
        .Text(", this is ")
        .Italic("emphasized")
        .Text(", and this is ")
        .Code("inline code")
        .Text(". See ")
        .Link("the docs", "https://example.com/docs")
        .Text("."));

Console.WriteLine(document.ToMarkdown());
```

### PowerShell

```powershell
New-OfficeMarkdown -Path .\ReleaseNotes.md {
    Add-OfficeMarkdownHeading -Level 1 -Text 'Release notes'
    Add-OfficeMarkdownParagraph {
        Add-OfficeMarkdownText 'Status: '
        Add-OfficeMarkdownText 'Healthy' -Bold
    }
}
```

## Add lists and task lists

### C#

```csharp
using OfficeIMO.Markdown;

string[] features = { "Cross-platform", "No dependencies", "Fluent API" };

MarkdownDocument document = MarkdownDocument.Create()
    .H2("Features")
    .Ul(features)
    .H2("Steps")
    .Ol(list => list
        .Item("Build")
        .Item("Test")
        .Item("Publish"))
    .H2("Checklist")
    .Ul(list => list
        .ItemTask("Packages published", done: true)
        .ItemTask("Change approved"))
    .H2("Terms")
    .Dl(list => list
        .Item("API", "Application Programming Interface")
        .Item("SDK", "Software Development Kit"));

Console.WriteLine(document.ToMarkdown());
```

### PowerShell

```powershell
New-OfficeMarkdown -Path .\Plan.md {
    Add-OfficeMarkdownList -Items 'Cross-platform', 'No dependencies', 'Fluent API'
    Add-OfficeMarkdownList -Items 'Build', 'Test', 'Publish' -Numbered
    Add-OfficeMarkdownTaskList -Items 'Packages published', 'Change approved' -Completed 0
    Add-OfficeMarkdownDefinitionList -Definition @{ API = 'Application Programming Interface'; SDK = 'Software Development Kit' }
}
```

## Add a table

Build rows by hand, or pass objects and let the builder read their properties. `TableAuto` and `MarkdownTableOptions` right-align numeric columns and center date columns.

### C#

```csharp
using OfficeIMO.Markdown;

Product[] products = {
    new("Widget A", 9.99m, 150),
    new("Widget B", 14.99m, 75)
};

MarkdownDocument document = MarkdownDocument.Create()
    .H2("Team")
    .Table(table => table
        .Headers("Name", "Role", "Location")
        .Row("Alice", "Developer", "New York")
        .Row("Bob", "Designer", "London")
        .Align(MarkdownColumnAlignment.Left, MarkdownColumnAlignment.Center, MarkdownColumnAlignment.Right))
    .H2("Stock");

TableBlock stock = document.AddTable(products, new MarkdownTableOptions { AlignNumeric = true });
stock.Rows.Add(new[] { "Widget C", "4.50", "300" });

document.H2("Selected columns");
document.AddTable(products,
    ("Product", p => p.Name),
    ("Price ($)", p => p.Price.ToString("F2")),
    ("In stock", p => p.Stock > 100 ? "Yes" : "Low"));

Console.WriteLine(document.ToMarkdown());

public sealed record Product(string Name, decimal Price, int Stock);
```

### PowerShell

```powershell
$rows = Get-Process | Select-Object -First 5 Name, Id, CPU
New-OfficeMarkdown -Path .\Processes.md {
    Add-OfficeMarkdownHeading -Level 2 -Text 'Top processes'
    Add-OfficeMarkdownTable -InputObject $rows
}

# Or get the table as Markdown text
$markdown = $rows | ConvertTo-OfficeMarkdown
```

## Add code blocks, quotes and callouts

Callouts are written as GitHub alerts (`> [!NOTE]`).

### C#

```csharp
using OfficeIMO.Markdown;

MarkdownDocument document = MarkdownDocument.Create()
    .Code("csharp", "Console.WriteLine(\"Hello, World!\");")
    .Code("json", "{ \"name\": \"OfficeIMO\" }")
    .Caption("package.json configuration")
    .Quote("The best way to predict the future is to invent it.")
    .Quote(q => q
        .Line("First line of the quote.")
        .Line("Second line of the quote."))
    .Callout("note", "Important", "This feature requires .NET 8 or later.")
    .Callout("warning", "Breaking change", "The API signature changed in 2.0.");

Console.WriteLine(document.ToMarkdown());
```

### PowerShell

```powershell
New-OfficeMarkdown -Path .\Cookbook.md {
    Add-OfficeMarkdownCode -Language 'powershell' -Text 'Get-Process | Select-Object -First 5'
    Add-OfficeMarkdownQuote -Text 'The release is ready when the checks agree.'
    Add-OfficeMarkdownCallout -Kind 'note' -Title 'Important' -Text 'This feature requires .NET 8 or later.'
}
```

## Add images, collapsible details and rules

Image sizes are `OfficeLength` values. Plain pixel numbers are written as `{width=… height=…}` hints that the HTML, Word and PDF converters honor.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Markdown;

MarkdownDocument document = MarkdownDocument.Create()
    .Image("screenshot.png", "Dashboard screenshot", "Dashboard")
    .Image("logo.png", "Logo", width: OfficeLength.Pixels(200), height: OfficeLength.Pixels(60))
    .Hr()
    .Details("Click to expand", body => body
        .P("This content is hidden by default.")
        .Code("sh", "echo 'Revealed!'"));

Console.WriteLine(document.ToMarkdown());
```

### PowerShell

```powershell
New-OfficeMarkdown -Path .\Report.md {
    Add-OfficeMarkdownImage -Path '.\Assets\architecture.png' -Alt 'Architecture diagram' -Width 640px
    Add-OfficeMarkdownHorizontalRule
    Add-OfficeMarkdownDetails -Summary 'Implementation notes' { Add-OfficeMarkdownParagraph -Text 'Hidden by default.' }
}
```

## Add front matter and a table of contents

Front matter is written as YAML. `Toc` inserts a table of contents built from the document's headings when it is written.

### C#

```csharp
using OfficeIMO.Markdown;

MarkdownDocument document = MarkdownDocument.Create()
    .FrontMatter(new {
        title = "Deployment checklist",
        date = "2026-01-15",
        tags = new[] { "runbook", "deployment" }
    })
    .H1("Deployment checklist")
    .Toc(options => {
        options.Title = "Contents";
        options.MinLevel = 2;
        options.MaxLevel = 3;
    })
    .H2("Pre-flight")
    .P("Check the packages.")
    .H2("Rollout")
    .P("Deploy to the first ring.");

document.Save("checklist.md");
```

### PowerShell

```powershell
New-OfficeMarkdown -Path .\Checklist.md {
    Add-OfficeMarkdownFrontMatter -InputObject @{ title = 'Deployment checklist'; tags = @('runbook', 'deployment') }
    Add-OfficeMarkdownHeading -Level 1 -Text 'Deployment checklist'
    Add-OfficeMarkdownTableOfContents -Title 'Contents' -MinLevel 2 -MaxLevel 3 -Placement Top
    Add-OfficeMarkdownHeading -Level 2 -Text 'Pre-flight'
    Add-OfficeMarkdownHeading -Level 2 -Text 'Rollout'
}
```
