---
title: "Redact literal text from a PDF"
description: "Find literal text matches, remove the matched PDF content, and verify extracted text, encoded strings, and decoded streams before download."
meta.workflow_id: "redact"
meta.eyebrow: "Secure a PDF"
meta.source_format: "PDF and literal text"
meta.destination_format: "Verified redacted PDF"
meta.package: "OfficeIMO.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_url: "/apps/officeimo-converter/?workspace=pdf&tool=redact"
meta.primary_label: "Redact text in the browser"
meta.secondary_url: "/docs/pdf/security/"
meta.secondary_label: "Read the security guide"
meta.summary_title: "Redaction summary"
meta.limit: "The browser route performs case-insensitive literal matching; image-only text requires an OCR-aware policy."
meta.related_url: "/pdf/"
meta.related_label: "Browse all PDF workflows"
meta.howto.name: "Remove and verify literal PDF text"
meta.howto.description: "Plan matches, confirm the destructive output change, apply redaction, and reject any failed verification."
meta.howto.steps:
  - name: "Search"
    text: "Choose one PDF and enter the exact literal text that must be removed."
  - name: "Confirm"
    text: "Review the destructive-action warning and require at least one match."
  - name: "Verify"
    text: "Apply the plan and download the PDF only after removal checks pass."
---

Redaction must remove content, not merely draw a black rectangle over visible text. The browser route plans case-insensitive literal matches, rewrites the document, and verifies concrete marker variants before offering the result.

## Redact from .NET

```csharp
using OfficeIMO.Pdf;

PdfDocument source = PdfDocument.Load("case-file.pdf");
var search = new PdfRedactionSearchOptions { MatchCase = false };
search.AddLiteral("Account 1234");

PdfRedactionPlan plan = source.Redactions.Search(search);
PdfDocument redacted = source.Redactions.Apply(plan);

var verify = new PdfRedactionVerificationOptions { MatchCase = false };
verify.RequireRemovedText("Account 1234");
PdfRedactionVerificationReport report = redacted.Redactions.Verify(verify);
report.ThrowIfFailed();

File.WriteAllBytes("case-file.redacted.pdf", redacted.ToBytes());
```

The verification report checks extracted text, raw PDF bytes, encoded strings, and decoded streams according to the selected markers. Keep the plan and report when the redaction decision must be auditable.

## Important limits

Literal search depends on readable PDF text. It does not discover text embedded only in scanned images, infer sensitive entities, or replace human review of surrounding context. OCR-assisted and pattern-based policies should remain separate, explicit workflows.
