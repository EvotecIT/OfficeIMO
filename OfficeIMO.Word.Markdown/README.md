# OfficeIMO.Word.Markdown - Word and Markdown conversion

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word.Markdown)](https://www.nuget.org/packages/OfficeIMO.Word.Markdown)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Word.Markdown?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word.Markdown)

`OfficeIMO.Word.Markdown` converts between `OfficeIMO.Word` documents and `OfficeIMO.Markdown` documents.

## Install

```powershell
dotnet add package OfficeIMO.Word.Markdown
```

## Quick start

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

using var document = WordDocument.Create();
document.AddParagraph("Hello");

string markdown = document.ToMarkdown();
using var fromMarkdown = "# Title\n\nBody".LoadFromMarkdown();
```

## AST-first conversion

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Word.Markdown;

MarkdownDoc markdownDocument = "<table><tr><td><p>Line 1</p><p>Line 2</p></td></tr></table>"
    .LoadFromHtml();

using var wordDocument = markdownDocument.ToWordDocument();
```

Use `MarkdownDoc.ToWordDocument()` when you already have a typed AST and want to avoid flattening back to Markdown text before Word conversion.

## What it maps

- Word to Markdown with headings, paragraphs, lists, task items, tables, images, links, code, footnotes, and GitHub-friendly output.
- Markdown to Word through the typed `OfficeIMO.Markdown` model.
- Markdown image layout options such as local-image allowance and page-content-width fitting.
- Selected AST-preserved inline HTML wrappers such as underline, superscript, and subscript into Word run formatting.

## Boundaries

- Word document modeling belongs in `OfficeIMO.Word`.
- Markdown parsing and AST behavior belongs in `OfficeIMO.Markdown`.
- HTML ingestion belongs in `OfficeIMO.Markdown.Html` or `OfficeIMO.Word.Html`, depending on the desired source model.
- PDF output belongs in `OfficeIMO.Word.Pdf` and `OfficeIMO.Pdf`.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)
