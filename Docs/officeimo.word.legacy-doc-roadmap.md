# DOC and DOCX compatibility

OfficeIMO.Word provides first-party, dependency-free support for Office Open XML `.docx` and the supported Word 97-2003 binary `.doc` subset. Microsoft Word, COM automation, LibreOffice, and third-party conversion libraries are not used at runtime.

This document is the current capability contract. It replaces the implementation roadmap.

## Normal API

Use the same `WordDocument` surface for both formats:

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("input.doc");
Console.WriteLine(document.SourceFormat); // WordFileFormat.Doc

document.Save("output.docx");
document.Save("copy.doc", new WordSaveOptions {
    LossPolicy = WordConversionLossPolicy.Allow
});

byte[] docx = document.ToBytes();
byte[] doc = document.ToDoc();
```

For an independent copy that does not change the current document association, use `SaveCopy`. For a writable stream, call `Save(stream, WordFileFormat.Docx)` or `Save(stream, WordFileFormat.Doc)`.

For a file-to-file conversion with a structured result:

```csharp
WordDocumentConversionResult result = WordDocument.Convert(
    "input.doc",
    "output.docx",
    new WordDocumentConversionOptions {
        FileConflictPolicy = WordConversionFileConflictPolicy.FailIfExists,
        LossPolicy = WordConversionLossPolicy.Block
    });

foreach (WordConversionDiagnostic diagnostic in result.Diagnostics) {
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
}
```

The defaults are intentionally conservative:

- content, not only the extension, determines the source format;
- same-format conversion is rejected;
- an existing destination is not replaced unless `Replace` is selected, and a read-only destination is never replaced;
- known conversion loss blocks conversion and normal saves;
- output is staged and committed atomically, so a failed save does not expose a partial file;
- cross-family OLE input, such as XLS passed to Word, is rejected with a format-specific error.

Set `LossPolicy = WordConversionLossPolicy.Allow` only after reviewing the reported legacy features. The policy is available on both `WordDocumentConversionOptions` and `WordSaveOptions`.

## DOC import capability

The DOC reader projects supported content into the normal OfficeIMO Word model. Current covered families include:

| Family | DOC to DOCX behavior |
|---|---|
| Paragraphs, runs, tabs, line/page/column breaks | Projected |
| Common character and paragraph formatting | Projected |
| Built-in and custom paragraph styles | Projected |
| Simple and supported nested tables | Projected |
| Sections, page setup, headers, and footers | Projected |
| Bookmarks and supported internal/external hyperlinks | Projected |
| Static and supported field display results | Projected |
| Footnotes and endnotes, including supported formatting | Projected |
| Comments with readable comment tables | Projected |
| Revision-tracking settings | Projected |
| Scalar core, application, and custom properties | Projected |
| Pictures, drawings, text boxes, and richer visual payloads | Diagnosed as preserve-only unless a supported projection exists |
| VBA, ActiveX, embedded packages, and OLE objects | Diagnosed as preserve-only |
| Damaged, encrypted, or unsupported binary structures | Rejected or diagnosed before output |

A readable feature is not automatically writable to DOC. DOCX can represent a broader model than the native DOC writer.

## Native DOC write capability

The native writer covers the tested binary subset, including paragraphs and runs, common formatting, styles, sections and page setup, supported headers and footers, simple tables and supported nesting, bookmarks, supported hyperlinks and static fields, footnotes and endnotes, and scalar document properties.

The writer preflights the complete document before committing output. Unsupported destination features—such as comments, tracked revision markup, images, drawings, embedded objects, unsupported content-control shapes, or richer table/story structures—raise `NotSupportedException` and leave an existing destination intact.

This is practical feature parity, not a claim that arbitrary DOCX packages can be represented in the older DOC format.

## Detailed import assessment

Normal application code can use a cached compact summary:

```csharp
using OfficeIMO.Word.LegacyDoc;

using LegacyDocLoadResult load = WordDocument.LoadLegacyDocWithReport("input.doc");
LegacyDocImportSummary summary = load.Summary;

if (summary.HasConversionLoss) {
    load.EnsureNoConversionLoss();
}
```

For corpus analysis or forensic detail, use `load.AdvancedDocument` and `load.CreateAdvancedImportReport()`. Import options use the common names `MaxInputBytes` and `ReportUnsupportedContent`. File conversion always enables unsupported-content discovery—even when a supplied import option disables reporting—because `LossPolicy.Block` must never be bypassed silently. Import options are selected from detected physical content, so limits and XLS passwords still apply when a legacy file has a misleading extension.

## Breaking API cleanup

The parity API intentionally uses one vocabulary:

| Removed API | Use |
|---|---|
| `SaveAs(path/stream)` | `SaveCopy(path/stream)` |
| `SaveAsByteArray()` | `ToBytes()` |
| `SaveAsMemoryStream()` | `ToStream()` |
| `WasLoadedFromLegacyDoc` | `SourceFormat == WordFileFormat.Doc` |
| `MaxWordDocumentStreamBytes` | `MaxInputBytes` |
| `ReportUnsupportedFeatures` | `ReportUnsupportedContent` |
| positional overwrite conversion flag | `FileConflictPolicy` |
| save-triggered application launch | Call `OpenInApplication(path)` explicitly after a successful save |
| lossy conversion Boolean | `LossPolicy` |

## Validation

The normal automated test lane is dependency-free. Optional desktop Word validation is explicitly skipped unless `OFFICEIMO_RUN_LEGACY_DOC_COM_VALIDATION` is enabled. When enabled, missing Windows, Word, or required corpus inputs fail the lane instead of producing a false pass.
