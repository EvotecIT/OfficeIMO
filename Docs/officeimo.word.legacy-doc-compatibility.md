# DOC and DOCX compatibility

OfficeIMO.Word provides first-party, dependency-free support for Office Open XML `.docx` and the supported Word 97-2003 binary `.doc` subset. Microsoft Word, COM automation, LibreOffice, and third-party conversion libraries are not used at runtime.

This document is the current capability contract.

## Normal API

Use the same `WordDocument` surface for both formats:

```csharp
using OfficeIMO.Word;

using WordDocument document = WordDocument.Load("input.doc");
Console.WriteLine(document.SourceFormat); // WordFileFormat.Doc

document.Save("output.docx");
document.Save("copy.doc", new WordSaveOptions {
    LossPolicy = OfficeConversionLossPolicy.Allow
});

byte[] docx = document.ToBytes();
byte[] doc = document.ToBytes(WordFileFormat.Doc);
```

For an independent copy that does not change the current document association, use `SaveCopy`. For a writable stream, call `Save(stream, WordFileFormat.Docx)` or `Save(stream, WordFileFormat.Doc)`.

For a file-to-file conversion with a structured result:

```csharp
WordDocumentConversionResult result = WordDocument.Convert(
    "input.doc",
    "output.docx",
    new WordDocumentConversionOptions {
        FileConflictPolicy = OfficeConversionFileConflictPolicy.FailIfExists,
        LossPolicy = OfficeConversionLossPolicy.Block
    });

foreach (OfficeConversionDiagnostic diagnostic in result.Diagnostics) {
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

Set `LossPolicy = OfficeConversionLossPolicy.Allow` only after reviewing the reported legacy features. The policy is available on both `WordDocumentConversionOptions` and `WordSaveOptions`.

## DOC import capability

The DOC reader projects supported content into the normal OfficeIMO Word model. Current covered families include:

| Family | DOC to DOCX behavior |
|---|---|
| Paragraphs, runs, tabs, line/page/column breaks | Projected |
| Common character and paragraph formatting | Projected |
| Built-in and custom paragraph styles | Projected |
| Supported list definitions, marker formatting, starts and instance restarts | Projected; missing, malformed or unsupported definitions report numbering loss |
| Simple and supported nested tables | Projected |
| Sections, page setup, headers, and footers | Projected, including supported tables in default, first-page, and even-page stories |
| Bookmarks and supported internal/external hyperlinks | Projected |
| Static and supported field display results | Projected |
| Footnotes and endnotes, including supported formatting | Projected |
| Comments with readable comment tables | Projected |
| Revision-tracking settings | Projected |
| Scalar core, application, and custom properties | Projected |
| Supported inline pictures and inset source crops | Projected as editable images; unsupported transforms and effects are diagnosed |
| Floating drawings, text boxes, and richer visual payloads | Diagnosed as preserve-only unless a supported projection exists |
| VBA, ActiveX, embedded packages, and OLE objects | Diagnosed as preserve-only |
| Damaged, encrypted, or unsupported binary structures | Rejected or diagnosed before output |

A readable feature is not automatically writable to DOC. DOCX can represent a broader model than the native DOC writer.

## Native DOC write capability

The native writer covers the tested binary subset, including paragraphs and runs, common formatting, styles, sections and page setup, supported headers and footers, simple tables and supported nesting, bookmarks, supported hyperlinks and static fields, footnotes and endnotes, and scalar document properties.

Default, first-page, and even-page headers and footers retain supported tables as editable rows and cells, alongside ordinary story paragraphs. Their tables use the same width, merge, border, palette-shading, nesting, and formatting limits as body tables. Hyperlinks and inline pictures resolve against the containing header or footer part, including pictures within supported nested tables.

Lists retain supported native numbering formats, marker text, alignment, suffixes, paragraph/run formatting, abstract starting values and per-instance start or formatting overrides. An explicit instance restart takes precedence over the abstract start. Native definitions contain level 0 alone or all nine levels; starts range from 0 through 32,767 and instance IDs from 1 through 32,767. Sparse instance IDs keep their paragraph references. A marker can contain at most one placeholder per available level: one at level 0, two at level 1, up to nine at level 8. Linked numbering styles, picture bullets, section-break restarts, unsupported level properties and references to absent instances or levels raise `NotSupportedException` before output is committed. Missing or malformed binary list tables produce a `Numbering` entry in `LegacyDocUnsupportedFeatures`; review that loss before explicitly allowing output.

List-tab alignment and positions remain editable when importing native numbering definitions, saving DOC, or projecting to DOCX. Native import reads Word's list-level tab additions in `sprmPChgTabs` when that operation has no range deletions. Paragraph, style and numbering tabs share the native alignment and leader mapping, including list tabs. Native output sorts tab positions and accepts up to 64 additions and 64 clears per tab operation, subject to the binary operand size limit. Larger tab sets raise `NotSupportedException`; save as DOCX to retain them. Native `sprmPChgTabs` operations containing range deletions are not projected. Retaining the tab metadata alone does not qualify numbering layout compatibility options in PDF output.

Supported nested tables retain their cell boundaries, multiple paragraphs per cell, widths and row/cell settings. Native output containing nested tables declares the Word 2000 binary format so Microsoft Word interprets the nested grid. Documents without nested tables retain the Word 97 format declaration.

Table layout retains each format's default. A DOC table without an AutoFit flag imports with `LayoutMode` set to `Fixed`. When saving DOCX content as DOC, an omitted layout setting retains DOCX's AutoFit behavior. `LayoutMode` reads direct settings, named table styles, their base styles and the document's default table style. Explicit and supported inherited layout settings take precedence over the format default. PDF conversion also retains the effective layout algorithm; preferred cell widths remain sizing inputs when layout defaults to AutoFit.

The writer also covers supported inline pictures and inset source crops. `OptimizeImages` can downsample or recompress their projected media before a native DOC save. Default fonts and Normal style formatting are materialized in the native stylesheet, including multiple, exact and at-least line spacing. Exact line spacing accepts 1–31,680 twips; exact zero is unrepresentable in DOC and is rejected before writing output. Paragraph boundaries are retained even when adjacent paragraphs have identical formatting.

Paragraph and style pagination flags retain explicit false values as well as true values. This includes page-break-before, keep-with-next, keep-lines-together, widow/orphan control, contextual spacing, line-number suppression, hyphenation suppression, and mirrored indentation. Direct outline levels retain values 0–9, including body text (9); built-in Heading styles retain their required heading levels. These are file-format preservation contracts; fixed-layout export has separate rendering limits.

Page setup retains section gutter widths and right-edge gutter settings. Document options retain the default tab interval, mirrored margins and top gutters alongside revision tracking and odd/even header selection.

Native DOC output retains footnote and endnote placement, numbering format, starting number and restart policy. The emitted Word 97 and Word 2000 formats store these options for the whole document. Sections with different effective note options raise `NotSupportedException`; save as DOCX to retain those section-specific settings. When importing later Word formats, the reader uses the effective FIB version and retains their authoritative section note settings.

Each section retains its own start type: continuous, next-column, next-page, odd-page, or even-page. Native section marks remain distinct from manual page breaks within paragraphs. An empty section-mark paragraph retains its paragraph and paragraph-mark formatting; import attaches the section to that paragraph without adding another terminator. `AddSection()` continues page numbering; explicit numbering restarts remain properties of the new section. PDF and image export resolve odd/even starts and numbering restarts through their layout engines.

The writer preflights the complete document before committing output. Unsupported destination features, including floating drawings, rotated or mirrored inline pictures, unsupported image effects, embedded objects, and richer table/story structures, raise `NotSupportedException` and leave an existing destination intact. The [generated capability matrix](Compatibility/generated/word-legacy-doc.md) lists the current feature families.

When an application needs a non-throwing gate before selecting a destination, run the real encoder without committing a file:

```csharp
LegacyDocWriteAssessment assessment = document.AssessLegacyDocWrite();
if (assessment.IsSupported) {
    document.Save("output.doc");
} else {
    Console.WriteLine($"{assessment.DiagnosticCode}: {assessment.Message}");
}
```

`AssessLegacyDocWrite()` intentionally executes the same native encoder as `Save`/`ToBytes`, so its answer cannot drift from a second hand-maintained feature checklist. It allocates the candidate DOC bytes and reports their encoded size, but does not commit an artifact.

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

For corpus analysis or forensic detail, use `load.AdvancedDocument` and `load.CreateAdvancedImportReport()`. Import options use the common names `MaxInputBytes` and `ReportUnsupportedContent`. File conversion always enables unsupported-content discovery—even when a supplied import option disables reporting—because `LossPolicy.Block` must never be bypassed silently. Import options are selected from detected physical content, so format-specific limits still apply when a legacy file has a misleading extension.

## Validation

The normal automated test lane is dependency-free. Optional desktop Word validation is explicitly skipped unless `OFFICEIMO_RUN_LEGACY_DOC_COM_VALIDATION` is enabled. When enabled, missing Windows, Word, or required corpus inputs fail the lane instead of producing a false pass.
