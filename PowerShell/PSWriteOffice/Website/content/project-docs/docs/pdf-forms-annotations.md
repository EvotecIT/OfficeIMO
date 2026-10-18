---
title: "Create and Process PDF Forms and Annotations"
description: "Create interactive form fields, inspect and fill existing forms, manage annotations, and flatten only when the delivery policy requires it."
layout: docs
---

PDF forms are interactive data structures, not just boxes painted on a page. Keep them interactive while recipients must enter or review values, then flatten a delivery copy only when editing must stop.

## Author fields in the PDF DSL

`PdfFormField` supports text, check box, choice, multi-select choice, and radio-button fields inside a generated document. Name every field predictably so later automation can inspect or fill it.

```powershell
PdfFormField -Name Reviewer -Type Text -Value 'Unassigned' -Width 320 -Height 24
PdfFormField -Name Decision -Type Choice -Options Approve,Reject,Defer -Value Defer
```

The [forms recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Pdf/Recipe-Pdf-Forms.ps1) creates text, choice, and check-box fields in one composition block.

## Fill, inspect, or flatten

Use `Set-OfficePdfForm` to fill existing fields and `Get-OfficePdfFormField` to read them back. `ConvertTo-OfficePdfFlatForm` makes field appearances part of page content. For annotations, use `Get-OfficePdfAnnotation`, `Set-OfficePdfAnnotation`, `Remove-OfficePdfAnnotation`, and `ConvertTo-OfficePdfFlatAnnotation` according to the review lifecycle.

Flattening is a one-way delivery decision for the output copy. Preserve the interactive source when the workflow may need another review round.
