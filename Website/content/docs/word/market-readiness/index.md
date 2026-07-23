---
title: "Production Word workflows"
description: "Build template-driven DOCX files, review document changes, convert HTML and Markdown, and generate inspectable validation proof with OfficeIMO.Word."
meta.seo_title: "Production DOCX automation workflows with OfficeIMO.Word"
order: 60
---

OfficeIMO.Word creates and edits `.docx` files without Microsoft Word or COM automation. These workflows combine the Word object model with templates, converters, structured comparison, diagnostics, and Open XML validation so an evaluator can inspect both the code and the resulting artifact.

## Generate documents from templates and data

Use merge fields, repeated table rows, repeated blocks, and content controls to turn one template into many documents. Template inspection can identify missing or ambiguous bindings before a batch is written.

Typical uses include:

- contracts and policy documents assembled from approved clauses
- invoices, order confirmations, and customer statements
- status reports and evidence packs built from application data
- forms that must be filled and read through content controls

Start with [templates and forms](/docs/word/templates/) for the binding APIs and validation flow.

## Compare documents and produce a redline

OfficeIMO.Word can compare document structure, generate a visible comparison document, and work with comments and supported revision metadata.

```csharp
WordComparisonResult result =
    WordDocumentComparer.CompareStructure(original, updated);

foreach (WordComparisonFinding finding in result.Findings) {
    Console.WriteLine($"{finding.ChangeKind}: {finding.DetailedLocation}");
}
```

Structured comparison covers paragraphs, tables, rows, cells, headers, footers, and images. Use the [review and comparison guide](/docs/word/review/) for comments, tracked changes, accepting or rejecting supported revisions, and redline output.

## Convert HTML and Markdown to DOCX

`OfficeIMO.Word.Html` and `OfficeIMO.Word.Markdown` convert web and Markdown content through dedicated packages. The proof workflow retains the source input, generated DOCX, round-trip output where meaningful, conversion diagnostics, and validation status.

Continue with [Word conversion and rendering](/docs/word/conversion/) for supported entry points and package selection.

## Run the proof gallery

The repository contains an executable gallery that produces:

- two HTML-to-DOCX scenarios
- two Markdown-to-DOCX scenarios
- a template assembly and form-binding scenario
- a review and structured-diff scenario
- source inputs, generated artifacts, diagnostics, validation results, and `proof-manifest.json`

```shell
dotnet run --project OfficeIMO.Examples -- --word-market-readiness
```

[Inspect the proof-gallery source](https://github.com/EvotecIT/OfficeIMO/tree/master/OfficeIMO.Examples/Word/MarketReadiness) before running it, or browse the [Word API reference](/api/word/) for the underlying types.

## Choose the right package

| Workflow | Package |
|---|---|
| Create, edit, inspect, compare, or validate DOCX | `OfficeIMO.Word` |
| Convert between Word and HTML | `OfficeIMO.Word.Html` |
| Convert between Word and Markdown | `OfficeIMO.Word.Markdown` |
| Render Word content to PDF | `OfficeIMO.Word.Pdf` with `OfficeIMO.Pdf` |

Each package keeps its own dependencies and API surface, so an application can install only the workflow it uses.
