---
title: Parse and Inspect Markdown Files
description: Load a Markdown file into a typed model, list headings and anchors, read and edit YAML or TOML front matter, extract tables and code blocks, and check parse diagnostics.
order: 42
search_terms:
  - parse markdown
  - read md file
  - markdown front matter
  - markdown headings
  - markdown tables to objects
  - markdown ast
packages:
  - OfficeIMO.Markdown
cmdlets:
  - Open-OfficeMarkdown
  - Get-OfficeMarkdownHeading
  - Get-OfficeMarkdownFrontMatter
  - Get-OfficeMarkdownTable
  - Get-OfficeMarkdownNode
  - Save-OfficeMarkdown
commands:
  - officeimo reader read
limits: Math syntax ($...$) is parsed only when MarkdownLoadOptions.Math is enabled; front matter that is not valid YAML or TOML stays as text with a diagnostic.
verified: 2026-11-14
---

`MarkdownDocument.Load` reads a file and `MarkdownDocument.Parse` reads a string. Both return the same typed model: headings with their anchors, blocks such as `TableBlock` and `CodeBlock`, inline formatting, front matter and source spans. Use the model to validate documentation, extract data from Markdown reports, or make targeted edits and save the file back.

## List headings and anchors

`GetHeadingInfos` returns every heading with its level, plain text and the anchor the HTML renderer will use.

### C#

```csharp
using OfficeIMO.Markdown;

MarkdownDocument document = MarkdownDocument.Parse("""
# User guide

## Install the package

## First steps
""");

foreach (MarkdownDocument.HeadingInfo heading in document.GetHeadingInfos()) {
    Console.WriteLine($"{new string(' ', (heading.Level - 1) * 2)}{heading.Text} -> #{heading.Anchor}");
}

MarkdownDocument.HeadingInfo? install = document.FindHeading("Install the package");
Console.WriteLine(install?.Anchor);
Console.WriteLine(document.HasHeadingAnchor("first-steps"));
```

### PowerShell

```powershell
Get-OfficeMarkdownHeading -Path .\README.md -MinLevel 2 |
    Select-Object Level, Text, Anchor
```

## Read and edit front matter

Front matter between `---` lines is YAML; between `+++` lines it is TOML. Unedited front matter is written back byte for byte, and an edit rewrites only the value you change.

### C#

```csharp
using OfficeIMO.Markdown;

File.WriteAllText("post.md", """
---
title: Draft title
author:
  name: Ada
draft: true
---
# Post
""");

MarkdownDocument document = MarkdownDocument.Load("post.md");
FrontMatterBlock header = document.DocumentHeader!;

string? title = header.FindEntry("title")?.Value.AsString();
string? author = header.Select("author.name")?.AsString();
bool draft = header.TryGetValue("draft", out bool isDraft) && isDraft;
Console.WriteLine($"{title} by {author}, draft: {draft}");

header.Yaml!.Set("title", "Final title");
document.Save();
```

### PowerShell

```powershell
Get-OfficeMarkdownFrontMatter -Path .\post.md | Select-Object Key, Value
Get-OfficeMarkdownFrontMatter -Path .\post.md -Key 'title'
```

## Extract tables and code blocks

`DescendantTables` and `DescendantsOfType<T>` walk the whole document, including content nested in lists, quotes and callouts.

### C#

```csharp
using OfficeIMO.Markdown;

MarkdownDocument document = MarkdownDocument.Parse("""
| Service | Status |
| --- | --- |
| API | Green |
| Jobs | Yellow |

~~~powershell
Get-Service
~~~
""");

foreach (TableBlock table in document.DescendantTables()) {
    Console.WriteLine(string.Join(" | ", table.Headers));
    foreach (IReadOnlyList<string> row in table.Rows) {
        Console.WriteLine(string.Join(" | ", row));
    }
}

foreach (CodeBlock code in document.DescendantsOfType<CodeBlock>()) {
    Console.WriteLine($"{code.Language}: {code.Content}");
}
```

### PowerShell

```powershell
Open-OfficeMarkdown -Path .\Report.md | Get-OfficeMarkdownTable -AsObject
Get-OfficeMarkdownNode -Path .\README.md -NodeType '*Code*'
```

## Check parse diagnostics

Use the `Load` or `Parse` overload with `out OfficeConversionReport` when diagnostics, such as invalid front matter, must travel with the document.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Markdown;

MarkdownDocument document = MarkdownDocument.Parse("""
---
title: [unclosed
---
# Notes
""", null, out OfficeConversionReport report);

foreach (OfficeDiagnostic diagnostic in report.Diagnostics) {
    Console.WriteLine($"{diagnostic.Severity} {diagnostic.Code}: {diagnostic.Message}");
}

Console.WriteLine(document.GetHeadingInfos().Count);
```

## Edit a parsed document and save it

Parsed documents use the same builder methods as new ones, so you can add missing sections and write the result.

### C#

```csharp
using OfficeIMO.Markdown;

File.WriteAllText("README.md", "# MyProject\n\nA data processing library.\n");

MarkdownDocument document = MarkdownDocument.Load("README.md");
if (!document.HasHeading("Install")) {
    document
        .H2("Install")
        .Code("sh", "dotnet add package MyProject");
}

document.Save("README.md");
Console.WriteLine(document.ToMarkdown());
```

### PowerShell

```powershell
$document = Open-OfficeMarkdown -Path .\README.md
$document | Add-OfficeMarkdownHeading -Level 2 -Text 'Install'
$document | Add-OfficeMarkdownCode -Language 'sh' -Text 'dotnet add package MyProject'
$document | Save-OfficeMarkdown
$document | Close-OfficeMarkdown
```

## Read Markdown as text or JSON for indexing

The Reader turns a Markdown file into normalized Markdown or JSON chunks for search pipelines.

### CLI

```sh
officeimo reader read README.md --format json --output README.json
```
