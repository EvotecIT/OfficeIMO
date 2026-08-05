---
title: "Generate DOCX from Templates in .NET"
description: "Bind C# models to ordinary DOCX placeholders, repeat Word blocks, add conditional content, images, and links, or use MERGEFIELD and content-control workflows."
layout: docs
meta.seo_title: "Generate DOCX from Word templates in .NET | OfficeIMO"
---

OfficeIMO.Word supports three template styles. Choose the one that matches who owns the document and how much structure the application needs.

| Template style | Use it when | Entry point |
|---|---|---|
| Plain `{{Name}}` placeholders | A designer owns a DOCX layout and application data should bind in one pass | `WordTemplate.Apply(...)` |
| Word `MERGEFIELD` fields | Existing mail-merge templates and Word field formatting are part of the contract | `WordMailMerge` |
| Tagged content controls | The document is a typed form with checkboxes, dates, choices, pictures, or repeating sections | `FillContentControlValues(...)` |

## Bind a model to ordinary placeholders

Write placeholders directly in Word:

```text
Service summary for {{Client.Name}}

{{#each Services}}
{{Name}} — {{Hours}} hours
{{#Priority}}
Priority delivery
{{/Priority}}
{{/each Services}}

Portal: {{Portal}}
Logo: {{Logo}}
```

Repeated and conditional markers must be placed on their own paragraphs. The content between them can contain normal paragraphs, styled runs, lists, tables, text boxes, and nested blocks. Scalar placeholders can be split across Word runs; the replacement inherits the formatting of the first run containing the marker.

The dictionary overload avoids reflection and is the recommended path for trimming and NativeAOT:

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("service-template.docx");

var values = new Dictionary<string, object?> {
    ["Client"] = new Dictionary<string, object?> {
        ["Name"] = "Northwind Traders"
    },
    ["Services"] = new object[] {
        new Dictionary<string, object?> {
            ["Name"] = "Assessment",
            ["Hours"] = 8,
            ["Priority"] = true
        },
        new Dictionary<string, object?> {
            ["Name"] = "Implementation",
            ["Hours"] = 24,
            ["Priority"] = false
        }
    },
    ["Portal"] = new WordTemplateHyperlink(
        "Open customer portal",
        new Uri("https://example.com/customer")),
    ["Logo"] = new WordTemplateImage(
        File.ReadAllBytes("logo.png"),
        "logo.png",
        width: 96,
        height: 32,
        description: "Company logo")
};

WordTemplateResult result = WordTemplate.Apply(document, values).EnsureComplete();
document.Save("service-summary.docx");
```

`WordTemplate.Apply(document, model)` also accepts a POCO or anonymous object. That overload reflects over public properties and is annotated accordingly; use dictionaries when the application publishes with trimming or NativeAOT.

## Template syntax

| Syntax | Meaning |
|---|---|
| `{{Name}}` | Scalar value; `null` becomes empty text |
| `{{Customer.Name}}` | Nested dictionary key or public property |
| `{{this}}` | Current scalar item inside a repeated block |
| `{{#each Items}}` … `{{/each Items}}` | Repeat any enclosed Word blocks for an enumerable value |
| `{{#Approved}}` … `{{/Approved}}` | Include enclosed content when the Boolean value is `true` |

Nested blocks inherit their parent scope. An item can therefore use both `{{Name}}` from the current item and `{{Client.Name}}` from the outer model.

## Detect incomplete output

Missing scalar values are preserved in the DOCX by default and reported through `WordTemplateResult`. This makes a bad binding visible instead of silently turning it into blank content.

```csharp
WordTemplateResult result = WordTemplate.Apply(document, values);

if (!result.IsComplete) {
    Console.Error.WriteLine(string.Join(", ", result.MissingValueNames));
}

result.EnsureComplete();
```

Set `WordTemplateOptions.RemoveMissingPlaceholders` only when blank output is an intentional application policy. Missing or non-Boolean block values and non-enumerable repeated values fail immediately because leaving half-expanded document structure would be ambiguous.

## Preflight classic mail merge templates

`WordMailMerge.PreflightTemplate` inspects Word `MERGEFIELD` fields and reports conditional blocks, repeating blocks, malformed markers, and requested names that are missing. Run preflight when a template is uploaded or promoted, not only after a production batch has failed.

```csharp
using var template = WordDocument.Load("contract-template.docx");
var report = WordMailMerge.PreflightTemplate(
    template,
    mergeFieldNames: new[] { "CustomerName", "ContractNumber", "EffectiveDate" });

if (!report.CanBindTemplate) {
    throw new InvalidOperationException(report.ToJson());
}
```

`WordMailMerge.Execute` replaces fields in an already loaded document. `ExecuteBatch` loads the template for every record and writes one output per value set. Table-row and grouped-row helpers repeat structured table regions without rebuilding layout in code.

## Content-control forms

`ExtractContentControlValues` reads tagged controls into a dictionary. `ValidateContentControlValues` reports missing, unused, duplicate, and invalid values. `FillContentControlValues` applies validated values to supported text, choice, date, checkbox, image, and repeating-section controls.

Choose one key policy—tag, alias, or a documented fallback order—and use it consistently across template creation, validation, and filling. Export validation reports as JSON or Markdown when template approval is part of an operational workflow.

## Delivery checklist

- Choose plain placeholders, `MERGEFIELD`, or content controls deliberately; do not mix them by accident.
- Validate required data before mutating the document.
- Generate into a new output path; keep the approved template immutable.
- Reopen the output and inspect expected sections, controls, and review metadata.
- Refresh fields and tables of contents where the target client requires it.
- Apply document protection or package signing only after content generation is complete.

## Executable evidence

The repository proof gallery generates a plain-placeholder template and its bound DOCX, records replacement and block counts, and validates every generated DOCX with the Open XML validator. It also keeps the existing `MERGEFIELD`, repeated-row, and content-control examples in the same scenario.

```shell
dotnet run --project OfficeIMO.Examples -- --word-market-readiness
```

[Inspect the template proof source](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Examples/Word/MarketReadiness/MarketReadinessProofGallery.TemplateBinding.cs) or continue with the [production Word workflow](/docs/word/market-readiness/).
