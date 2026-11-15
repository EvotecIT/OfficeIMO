---
title: Markdown in .NET
description: "Create, parse, edit and render Markdown files in .NET with OfficeIMO.Markdown, plus the PSWriteOffice Markdown cmdlets for scripts and jobs."
order: 40
search_terms:
  - markdown
  - md file
  - markdown parser
  - markdown builder
  - markdown to html
packages:
  - OfficeIMO.Markdown
cmdlets:
  - New-OfficeMarkdown
  - Open-OfficeMarkdown
  - Get-OfficeMarkdownHeading
limits: OfficeIMO.Markdown does not execute raw HTML or scripts; raw HTML blocks are kept as text and handled by the HTML renderer's raw-HTML policy.
verified: 2026-11-14
---

`OfficeIMO.Markdown` builds, parses and writes Markdown through one typed model, `MarkdownDocument`. The same object renders HTML fragments or full pages, and adapter packages turn it into PDF (`OfficeIMO.Markdown.Pdf`) or Word (`OfficeIMO.Word.Markdown`). The package has no external dependencies and targets `netstandard2.0`, `net8.0` and `net10.0`.

```sh
dotnet add package OfficeIMO.Markdown
```

## Create a Markdown file

Chain the builder methods and save. Each shortcut (`H1`, `P`, `Ul`, `Table`, `Code`) appends a block and returns the document.

### C#

```csharp
using OfficeIMO.Markdown;

MarkdownDocument document = MarkdownDocument.Create()
    .H1("Project README")
    .P("A short description of the project.")
    .H2("Installation")
    .Code("sh", "dotnet add package MyProject")
    .H2("Usage")
    .Ul(list => list
        .Item("Import the namespace")
        .Item("Create a client"));

document.Save("README.md");
Console.WriteLine(document.ToMarkdown());
```

### PowerShell

```powershell
New-OfficeMarkdown -Path .\README.md {
    Add-OfficeMarkdownHeading -Level 1 -Text 'Project README'
    Add-OfficeMarkdownParagraph -Text 'A short description of the project.'
    Add-OfficeMarkdownHeading -Level 2 -Text 'Installation'
    Add-OfficeMarkdownCode -Language 'sh' -Text 'dotnet add package MyProject'
}
```

## Parse a Markdown file

`MarkdownDocument.Load` and `MarkdownDocument.Parse` return the typed model with headings, blocks, inlines, anchors and front matter.

### C#

```csharp
using OfficeIMO.Markdown;

MarkdownDocument document = MarkdownDocument.Parse("""
# Guide

## Install

## Usage
""");

foreach (MarkdownDocument.HeadingInfo heading in document.GetHeadingInfos()) {
    Console.WriteLine($"H{heading.Level}: {heading.Text} (#{heading.Anchor})");
}

Console.WriteLine(document.HasHeading("Install"));
```

### PowerShell

```powershell
Open-OfficeMarkdown -Path .\README.md | Get-OfficeMarkdownHeading -MinLevel 2
```

## Render Markdown to HTML

`ToHtml` returns a fragment or a complete page. Without a style the output is plain markup; choose `Clean`, `Word` or a GitHub style to add a stylesheet.

### C#

```csharp
using OfficeIMO.Markdown;

MarkdownDocument document = MarkdownDocument.Create()
    .H1("Hello")
    .P("World");

string fragment = document.ToHtml(new MarkdownHtmlExportOptions {
    Kind = MarkdownHtmlKind.Fragment
});

document.SaveAsHtml("hello.html", new MarkdownHtmlExportOptions {
    Kind = MarkdownHtmlKind.Document,
    Style = MarkdownHtmlStyle.GithubAuto,
    InjectTocAtTop = true
});

Console.WriteLine(fragment);
```

## Start from a scaffold

Scaffolds return a ready-made `MarkdownDocument` for common repository files, which you can extend before saving.

### C#

```csharp
using OfficeIMO.Markdown;

MarkdownDocument readme = MarkdownScaffold.Readme("MyProject", options => options
    .Description("A cross-platform library for data processing.")
    .NuGet("MyProject")
    .GettingStarted("dotnet add package MyProject", "var client = new MyClient();")
    .LicenseMIT());
readme.Save("README.md");

MarkdownChangelogScaffold.Changelog("MyProject").Save("CHANGELOG.md");
MarkdownContributingScaffold.Contributing("MyProject").Save("CONTRIBUTING.md");
```
