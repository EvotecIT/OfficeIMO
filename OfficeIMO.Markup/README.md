# OfficeIMO.Markup - semantic Office authoring

`OfficeIMO.Markup` is a Markdown-inspired authoring layer for OfficeIMO. It parses `.omd` or Markdown-like text into a semantic AST that can be validated, emitted as starter code, or exported to real Office files by target-specific renderers.

This project is built from the OfficeIMO source tree and is not published as a standalone NuGet package.

## Authoring model

The pipeline is intentionally staged:

- Parse Markdown, front matter, and OfficeIMO directives.
- Produce a semantic AST independent from C# and PowerShell APIs.
- Validate the AST against a document, workbook, or presentation profile.
- Emit starter C# or PowerShell code when a user wants to take over in code.
- Export real Office files through profile-specific exporter packages.

## Example

```markdown
---
profile: presentation
title: Quarterly Review
theme: evotec-modern
---

# Quarterly Review

@slide {
  layout: title-and-content
  transition: fade
}

- Revenue grew
- Churn improved

::notes
Open with the top-line result.
```

## Examples

### Parse and validate semantic markup

```csharp
using OfficeIMO.Markup;

OfficeMarkupParseResult result = OfficeMarkupParser.Parse(File.ReadAllText("quarterly-review.omd"));

foreach (OfficeMarkupDiagnostic diagnostic in result.Diagnostics) {
    Console.WriteLine($"{diagnostic.Severity}: {diagnostic.Message}");
}

if (result.HasErrors) {
    throw new InvalidOperationException("Markup contains validation errors.");
}

Console.WriteLine(result.Document.Profile);
Console.WriteLine(result.Document.Metadata["title"]);
```

### Emit starter C# for handoff

```csharp
using OfficeIMO.Markup;

OfficeMarkupParseResult result = OfficeMarkupParser.Parse(File.ReadAllText("document.omd"));

string csharp = new OfficeMarkupCSharpEmitter().Emit(result.Document,
    new OfficeMarkupEmitterOptions {
        FilePathVariable = "outputPath",
        IncludeHeader = true
    });

File.WriteAllText("document.generated.cs", csharp);
```

### Export the same semantic document to an Office file

```csharp
using OfficeIMO.Markup;
using OfficeIMO.Markup.Word;

OfficeMarkupParseResult result = OfficeMarkupParser.Parse(File.ReadAllText("status-brief.omd"));

result.Document.SaveAsWord("status-brief.docx", new MarkupToWordOptions {
    });
```

## What it understands

- Common Markdown structure through `OfficeIMO.Markdown`: headings, paragraphs, lists, fenced code, images, and pipe tables.
- Front matter for document-level metadata and target profile selection.
- Container directives such as `@slide`, `@section`, and `@sheet`.
- Office-aware blocks such as `::notes`, `::chart`, `::mermaid`, `::range`, `::formula`, `::table`, `::textbox`, `::columns`, `::column`, and `::card`.
- Semantic chart, workbook, slide, document, layout, and style attributes that exporters can map to editable Office features.

## Boundaries

- `OfficeIMO.Markup` owns parsing, semantic AST, validation, and code emission.
- Word export belongs in `OfficeIMO.Markup.Word`.
- Excel export belongs in `OfficeIMO.Markup.Excel`.
- PowerPoint export belongs in `OfficeIMO.Markup.PowerPoint`.
- CLI workflows belong in `OfficeIMO.Markup.Cli`.
- VS Code authoring support belongs in `OfficeIMO.Markup.VSCode`.

## Related packages

| Package | Use it for |
| --- | --- |
| [OfficeIMO.Markup.Word](../OfficeIMO.Markup.Word/README.md) | Export markup documents to Word. |
| [OfficeIMO.Markup.Excel](../OfficeIMO.Markup.Excel/README.md) | Export markup workbooks to Excel. |
| [OfficeIMO.Markup.PowerPoint](../OfficeIMO.Markup.PowerPoint/README.md) | Export markup presentations to PowerPoint. |
| [OfficeIMO.Markup.Cli](../OfficeIMO.Markup.Cli/README.md) | Parse, validate, emit, and export from the command line. |
| [OfficeIMO.Markdown](../OfficeIMO.Markdown/README.md) | Markdown model used by the markup parser. |

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** None.
- **OfficeIMO:** `OfficeIMO.Markdown` and `OfficeIMO.Drawing`; the semantic authoring model and validation are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
