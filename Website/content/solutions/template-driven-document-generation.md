---
title: "Template-Driven Word Document Generation"
description: "Generate contracts, invoices, reports, and customer documents from approved DOCX templates in .NET, with explicit binding diagnostics and validation evidence."
layout: solution
meta.eyebrow: "Word automation"
meta.outcome: "Keep layout in Word and data in your application"
meta.primary_label: "Read the template guide"
meta.primary_url: "/docs/word/templates/"
meta.social_card_badge: "DOCX templates for .NET"
---

When business owners already maintain a Word layout, rebuilding that layout from paragraphs and tables in code creates two sources of truth. OfficeIMO.Word can keep the approved DOCX as the layout and bind application data into ordinary placeholders when a document is generated.

## Where this fits

This workflow is useful for:

- contracts and engagement letters with approved wording and branding
- invoices, statements, and order confirmations with repeated line items
- audit reports and evidence packs assembled from structured findings
- customer notices with optional clauses and account-specific links
- scheduled reports generated on Linux, macOS, or Windows without Word automation

The application owns data selection, authorization, storage, and delivery. The template owns typography, spacing, tables, headers, footers, and other Word layout decisions.

## A production-shaped pipeline

1. Store an immutable, reviewed DOCX template.
2. Validate the required model before opening the output document.
3. Copy or load the template into a new destination.
4. Bind the model with `WordTemplate.Apply(...)` and require a complete result.
5. Reopen and validate the generated DOCX.
6. Convert to PDF only when the delivery workflow needs a fixed-layout artifact.
7. Apply document protection or package signing after content generation is complete.

```csharp
using OfficeIMO.Word;

using WordDocument output = WordDocument.Load("contract-template.docx");
WordTemplateResult binding = WordTemplate.Apply(output, contractValues);
binding.EnsureComplete();
output.Save("contract-1042.docx");
```

Repeated marker regions can contain tables and lists, so line items and clause groups do not need separate layout code. Boolean blocks let the document include approved optional sections without concatenating WordprocessingML or HTML strings.

## Choose the right binding contract

Use the AOT-safe dictionary overload when model keys are part of an explicit document contract. Use the POCO overload for conventional applications where reflection is acceptable. Use classic `MERGEFIELD` templates when existing Word field behavior matters, and use content controls for typed form fields such as checkboxes, choices, dates, or picture controls.

## Evidence, not a success flag

`WordTemplateResult` records discovered and replaced placeholders, generated repeated blocks, evaluated conditions, and missing values. The repository proof gallery generates both the source template and bound output, then validates the DOCX package. This proves structural output and binding completeness; it does not replace visual approval in the Word viewers your users rely on.

Start with the [DOCX template guide](/docs/word/templates/), inspect the [executable proof source](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Examples/Word/MarketReadiness/MarketReadinessProofGallery.TemplateBinding.cs), and use [Word-to-PDF guidance](/convert/word-to-pdf/) when the final artifact must be PDF.
