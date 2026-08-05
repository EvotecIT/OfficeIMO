---
title: "Generate a DOCX from a Word Template in .NET"
description: "Use placeholders, repeated blocks, conditions, images, and links to generate DOCX files from application data without Microsoft Word or COM automation."
date: 2026-08-05
tags: [word, docx, dotnet, automation]
categories: [Tutorial]
author: "Przemyslaw Klys"
meta.seo_title: "Generate DOCX from a Word template in .NET without Word"
meta.howto.name: "Generate a DOCX from a Word template with OfficeIMO"
meta.howto.description: "Create placeholders in a DOCX, bind application data, require complete output, and save a new document."
meta.howto.steps:
  - name: "Prepare the DOCX"
    text: "Add scalar placeholders and put repeated or conditional markers on their own paragraphs."
  - name: "Build the model"
    text: "Map application data to an AOT-safe dictionary or a public-property object."
  - name: "Bind"
    text: "Call WordTemplate.Apply and inspect the result."
  - name: "Validate and save"
    text: "Require complete binding, save to a new path, and validate the generated artifact."
---

Generating a Word document from code does not have to mean recreating every margin, table, and branded heading in C#. When the layout already exists as an approved DOCX, keep it there and bind only the changing data.

OfficeIMO.Word supports ordinary `{{Name}}` placeholders, repeated Word blocks, Boolean sections, embedded images, and external hyperlinks. It runs in-process and does not require Microsoft Word or COM automation.

## 1. Prepare the Word template

Type scalar placeholders where values should appear:

```text
Service summary for {{Client.Name}}
```

Put repeated and conditional markers on their own paragraphs. The content between them can be a paragraph, list, table, or a combination of Word blocks:

```text
{{#each Services}}
{{Name}} — {{Hours}} hours
{{#Priority}}
Priority delivery
{{/Priority}}
{{/each Services}}
```

Word may split visible text across several formatting runs. The binder recognizes a placeholder across those runs and uses the first marker run as the replacement formatting source.

## 2. Bind application data

The dictionary overload is explicit and safe for trimmed or NativeAOT applications:

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

WordTemplateResult result = WordTemplate.Apply(document, values);
result.EnsureComplete();
document.Save("service-summary.docx");
```

A POCO overload is available when reflection over public properties is acceptable. The API marks that overload for trimming analysis instead of pretending arbitrary reflection is NativeAOT safe.

## 3. Treat missing values as a document defect

By default, an unresolved scalar marker remains visible and its name appears in `MissingValueNames`. That is useful in batch jobs: a missing customer field should not quietly become a blank contract.

```csharp
WordTemplateResult result = WordTemplate.Apply(document, values);

if (!result.IsComplete) {
    throw new InvalidOperationException(
        "Template data is incomplete: " + string.Join(", ", result.MissingValueNames));
}
```

Missing repeated collections and conditions fail immediately because partially expanding document structure would produce a misleading result.

## 4. Validate the artifact users receive

A successful method call proves that the binding pass completed. It does not prove that a font available on one machine exists in production, or that every target Word viewer lays out the document identically.

For production templates:

- reopen the generated DOCX and validate its Open XML package
- inspect representative outputs in the viewers your users rely on
- keep the approved template immutable and write to a new destination
- capture binding counts and missing-value diagnostics in job evidence
- validate PDF output separately if PDF is part of delivery

The OfficeIMO repository includes an executable proof scenario that writes the template, bound DOCX, binding-result summary, and Open XML validation results:

```shell
dotnet run --project OfficeIMO.Examples -- --word-market-readiness
```

[Read the complete template guide](/docs/word/templates/), review the [template-driven generation use case](/solutions/template-driven-document-generation/), or inspect the [proof source](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Examples/Word/MarketReadiness/MarketReadinessProofGallery.TemplateBinding.cs).
