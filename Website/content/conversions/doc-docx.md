---
title: "Convert DOC and DOCX in .NET"
description: "Read Word 97–2003 DOC files, convert them to DOCX, or deliver supported DOC output from .NET with explicit fidelity and preservation diagnostics."
meta.eyebrow: "Word conversion"
meta.source_format: "DOC"
meta.destination_format: "DOCX"
meta.package: "OfficeIMO.Word"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Word"
meta.runtime: ".NET on Windows, Linux, and macOS"
meta.howto.name: "Convert a DOC file to DOCX with OfficeIMO.Word"
meta.howto.description: "Analyze the source, choose an accepted compatibility policy, and create a modern DOCX file without Microsoft Word."
meta.howto.steps:
  - "Install|Add the OfficeIMO.Word NuGet package to the application."
  - "Analyze|Inspect the conversion report when the workflow has fidelity requirements."
  - "Convert|Call WordDocument.Convert with the source DOC and destination DOCX paths."
  - "Review|Record or reject findings according to the application's document policy."
---

Use `OfficeIMO.Word` when a .NET application needs to bring Word 97–2003 files into a modern document workflow without Microsoft Word or COM automation. The normal Word engine detects the source format, projects supported DOC content into the OfficeIMO object model, and records features that need approximation, visual fallback, preservation, or rejection.

## DOC to DOCX

```csharp
using OfficeIMO.Word;

WordDocumentConversionReport preview =
    WordDocument.AnalyzeConversion("archive.doc", "archive.docx");

WordDocumentConversionResult result =
    WordDocument.Convert("archive.doc", "archive.docx");
```

For an unattended archive job, use the preview report to decide whether the application accepts every finding before it creates the destination. Keep the report beside migration logs when auditability matters.

## DOCX to DOC

The same API supports the reverse direction for the documented native DOC subset:

```csharp
using OfficeIMO.Word;

WordDocumentConversionResult result =
    WordDocument.Convert("contract.docx", "contract.doc");
```

DOC has a smaller feature vocabulary than DOCX. Modern charts, content controls, custom XML, advanced drawing, or other features may require a selected fallback or may be blocked. OfficeIMO does not label a lossy result as fully native.

## When to use each direction

Choose DOC to DOCX for archive modernization, search ingestion, future editing, and integration with modern Office workflows. Choose DOCX to DOC only when a downstream system genuinely requires the legacy format and the compatibility report meets that system's needs.

For the current feature-level evidence, see [Word compatibility](/compatibility/#word). For authoring and review APIs after conversion, use the [Word guide](/docs/word/).
