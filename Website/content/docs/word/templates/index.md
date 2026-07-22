---
title: "Templates, Forms, and Mail Merge"
description: "Preflight DOCX templates, bind merge fields and content controls, repeat rows and blocks, and validate output before delivery."
layout: docs
---

OfficeIMO.Word supports both classic `MERGEFIELD` templates and structured content-control forms. Use merge fields for familiar mail-merge documents; use tagged content controls when the template is a form with typed values, choices, dates, checkboxes, images, or repeating sections.

## Preflight before generating output

`WordMailMerge.PreflightTemplate` inspects the template and reports merge fields, conditional blocks, repeating blocks, malformed markers, and requested names that are missing. Run preflight when a template is uploaded or promoted, not only after a production batch has failed.

```csharp
using OfficeIMO.Word;

using var template = WordDocument.Load("contract-template.docx");
var report = WordMailMerge.PreflightTemplate(
    template,
    mergeFieldNames: new[] { "CustomerName", "ContractNumber", "EffectiveDate" });

if (!report.CanBindTemplate) {
    throw new InvalidOperationException(report.ToJson());
}
```

## Merge one document or a batch

`WordMailMerge.Execute` replaces fields in an already loaded document. `ExecuteBatch` loads the template for every record and writes one output per value set. Table-row and grouped-row helpers repeat structured table regions without rebuilding the template layout in code.

Conditional markers and repeating blocks let the template own layout while application data owns values. Keep marker names stable and treat the preflight report as the template contract.

## Content-control forms

`ExtractContentControlValues` reads tagged controls into a dictionary. `ValidateContentControlValues` reports missing, unused, duplicate, and invalid values. `FillContentControlValues` applies validated values to supported text, choice, date, checkbox, image, and repeating-section controls.

Choose one key policy—tag, alias, or a documented fallback order—and use it consistently across template creation, validation, and filling. Export validation reports as JSON or Markdown when template approval is part of an operational workflow.

## Delivery checklist

- Preflight template fields, markers, and content controls.
- Validate required data before mutating the document.
- Generate into a new output path; keep the approved template immutable.
- Refresh fields and tables of contents where the target client requires it.
- Reopen the output and inspect expected sections, controls, and review metadata.
- Apply document protection or package signing only after content generation is complete.
