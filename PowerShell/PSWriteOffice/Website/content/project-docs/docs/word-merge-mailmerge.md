---
title: "Merge Word Documents and Generate Letters"
description: "Combine complete DOCX files or fill Word merge fields from PowerShell data for packs, letters, and repeatable document batches."
layout: docs
---

Word has two distinct merge workflows. Document merge appends complete files into a pack. Mail merge replaces named fields with values for one or many recipients.

## Combine complete documents

```powershell
Join-OfficeWordDocument `
    -Path '.\Cover.docx' `
    -AppendPath '.\Report.docx','.\Appendix.docx' `
    -OutputPath '.\Delivery-Pack.docx'
```

The [merge-documents recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Word/Recipe-Word-MergeDocuments.ps1) creates a cover, detail document, and appendix, then combines them in order.

## Fill merge fields

Add `MergeField` fields while authoring a template, then call `Invoke-OfficeWordMailMerge -Values` with a hashtable. For a batch, create or copy one output per record so each file remains independently reviewable.

```powershell
WordParagraph {
    WordText 'Hello '
    WordField -Type MergeField -Parameters '"FirstName"'
}
Invoke-OfficeWordMailMerge -Values @{ FirstName = 'Ada' }
```

The [mail-merge letters recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Word/Recipe-Word-MailMergeLetters.ps1) generates two personalized order confirmations from ordinary PowerShell hashtables.

## Which merge should you use?

Use document merge when the input files are already complete deliverables. Use mail merge when one layout is repeated over data. Use both when a generated letter must be appended to a standard terms document or evidence pack.
