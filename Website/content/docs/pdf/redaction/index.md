---
title: "Redact PDF Content"
description: "Find sensitive text with literal or regex searches, review the plan, remove the content and verify the redacted PDF with OfficeIMO.Pdf."
layout: docs
packages:
  - OfficeIMO.Pdf
cmdlets:
  - Remove-OfficePdfContent
limits: "Scanned pages need a text layer (OCR) before text can be found."
search_terms:
  - redact pdf
  - remove text pdf
  - black out
commands:
  - officeimo pdf redact plan
  - officeimo pdf redact apply
  - officeimo pdf redact verify
verified: 2026-11-14
---

Redaction removes content from the page: the matched text, and optionally the image and vector content underneath it, is deleted from the content stream and the area is painted over. Drawing a black rectangle on top of text is not redaction, because the text stays extractable.

OfficeIMO.Pdf redacts in three steps: **search** builds a reviewable plan, **apply** removes the planned content, and **verification** checks the saved bytes for anything that is still extractable.

## Redact text by literal or pattern

### C#

```csharp
using OfficeIMO.Pdf;

PdfDocument source = PdfDocument.Load("statement.pdf");

PdfRedactionPlan plan = source.Redactions.Search(new PdfRedactionSearchOptions {
    TextSelection = PdfRedactionTextSelection.MatchedGlyphs,
    ContentScope = PdfRedactionContentScope.TextOnly
}.AddLiteral("Contoso Ltd").AddRegex(@"\b\d{3}-\d{2}-\d{4}\b"));

Console.WriteLine($"{plan.Matches.Count} matches, reviewable: {plan.IsReviewable}");

PdfRedactionApplyResult result = source.Redactions.ApplyWithEvidence(plan).ThrowIfUnverified();
result.ToDocument().Save("statement-redacted.pdf");
```

`ThrowIfUnverified` throws when the evidence reports residual or inconclusive matches, so an unverified file is never saved by accident. `MatchedGlyphs` removes only the matched characters; the default removes the whole text block that contains a match.

### PowerShell

```powershell
Get-OfficePdfRedactionPlan -Path .\Contracts\MSA-Contoso.pdf -Pattern 'Account: \d+'
Remove-OfficePdfContent -Path .\Contracts\MSA-Contoso.pdf -OutputPath .\Contracts\MSA-Redacted.pdf -Text 'Contoso Ltd'
```

### CLI

The CLI separates planning, reviewed decisions, and application, and writes JSON evidence for each step:

```sh
officeimo pdf redact plan contract.pdf --recipe redaction.recipe.json --evidence contract.plan.json
officeimo pdf redact apply contract.pdf --recipe redaction.recipe.json --decisions contract.decisions.json --output contract-redacted.pdf --evidence contract-redacted.evidence.json
officeimo pdf redact verify contract.pdf --recipe redaction.recipe.json --decisions contract.decisions.json --output contract-redacted.pdf --evidence contract-verify.evidence.json
```

The recipe and decision file formats are described in [PDF workflows in the CLI](/docs/cli/pdf/).

## Limit the search

Restrict the search to specific pages, form fields, or case-sensitive matches:

```csharp
using OfficeIMO.Pdf;

var options = new PdfRedactionSearchOptions { MatchCase = true }.AddLiteral("Internal only");
options.PageNumbers.Add(1);
options.PageNumbers.Add(2);

PdfRedactionPlan plan = PdfDocument.Load("memo.pdf").Redactions.Search(options);
```
