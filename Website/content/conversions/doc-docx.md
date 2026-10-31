---
title: "Convert DOC and DOCX in .NET"
description: "Read Word 97–2003 DOC files, convert them to DOCX, or deliver supported DOC output from .NET with explicit fidelity and preservation diagnostics."
meta.eyebrow: "Word conversion"
meta.source_format: "DOC"
meta.destination_format: "DOCX"
meta.package: "OfficeIMO"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO"
meta.runtime: ".NET on Windows, Linux, and macOS"
meta.primary_label: "Get OfficeIMO"
meta.primary_url: "https://www.nuget.org/packages/OfficeIMO"
meta.secondary_label: "Check Word compatibility"
meta.secondary_url: "/compatibility/#word"
meta.summary_title: "DOC and DOCX summary"
meta.limit: "DOC and DOCX do not share every feature. Conversion reports approximated, preserved, dropped, or blocked content instead of hiding loss."
meta.related_label: "Explore all conversion guides"
meta.related_url: "/convert/guides/"
meta.howto.name: "Convert a DOC file to DOCX with OfficeIMO"
meta.howto.description: "Analyze the source, choose an accepted compatibility policy, and create a modern DOCX file without Microsoft Word."
meta.howto.steps:
  - name: "Install"
    text: "Add the OfficeIMO NuGet package to the application."
  - name: "Analyze"
    text: "Inspect the conversion report when the workflow has fidelity requirements."
  - name: "Convert"
    text: "Call OfficeConverter.Convert with the source DOC and destination DOCX paths."
  - name: "Review"
    text: "Record or reject findings according to the application's document policy."
---

Use `OfficeIMO` for file-to-file conversion of Word 97–2003 documents without Microsoft Word or COM automation. `OfficeIMO.Word` owns the Word object model and its native DOC support. The normal Word engine detects the source format, projects supported DOC content into the OfficeIMO object model, and records features that need approximation, visual fallback, preservation, or rejection.

## DOC to DOCX

```csharp
using OfficeIMO;

OfficeConversionReport preview =
    OfficeConverter.Analyze("archive.doc", "archive.docx");

OfficeConversionReport result =
    OfficeConverter.Convert("archive.doc", "archive.docx");
```

For an unattended archive job, use the preview report to decide whether the application accepts every finding before it creates the destination. Keep the report beside migration logs when auditability matters.

## DOCX to DOC

The same API supports the reverse direction for the documented native DOC subset:

```csharp
using OfficeIMO;

OfficeConversionReport result =
    OfficeConverter.Convert("contract.docx", "contract.doc");
```

DOC has a smaller feature vocabulary than DOCX. Modern charts, content controls, custom XML, advanced drawing, or other features may require a selected fallback or may be blocked. OfficeIMO does not label a lossy result as fully native.

## When to use each direction

Choose DOC to DOCX for archive modernization, search ingestion, continued editing, and integration with modern Office workflows. Choose DOCX to DOC only when a downstream system genuinely requires the legacy format and the compatibility report meets that system's needs.

For the current feature-level evidence, see [Word compatibility](/compatibility/#word). For authoring and review APIs after conversion, use the [Word guide](/docs/word/).
