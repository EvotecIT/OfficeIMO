# OfficeIMO.Word.Rtf

`OfficeIMO.Word.Rtf` maps RTF documents to and from `OfficeIMO.Word`. It also exposes RTF-facing document workflows that reuse the Word engines and return a combined fidelity report.

## Convert with diagnostics

```csharp
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;

RtfDocument rtf = RtfDocument.Load("input.rtf");
RtfConversionResult<WordDocument> conversion = rtf.ToWordDocumentResult();

using WordDocument word = conversion.Value;
conversion.Report.RequireNoLoss();
word.Save("output.docx");
```

The reverse API is `word.ToRtfDocumentResult()`. Compatibility helpers such as `ToWordDocument()`, `ToRtfDocument()`, `LoadFromRtf()`, and `ToRtf()` remain available when the caller does not need the report.

Supported text, tabs, breaks, images and note references retain their order when combined within one Word run. Conversion keeps surrounding text and each note reference without changing the Word source.

## Run Word workflows from RTF

```csharp
RtfWordWorkflowResult<int> result = rtf.FindAndReplaceResult(
    "Contoso Ltd.",
    "Contoso Europe");

Console.WriteLine($"Replacements: {result.WorkflowResult}");
result.RequireNoLoss();
result.Document.Save("updated.rtf");
```

Result-bearing workflows include:

- `MailMergeResult`
- `FindAndReplaceResult`
- `UpdateFieldsResult`
- `MergeResult`
- `CompareResult`

Each result contains the normalized RTF document, the workflow-specific result, and one `RtfConversionReport` covering input mapping, the Word operation, and output mapping. Mail merge, field evaluation, comparison, and Word document merge remain owned by `OfficeIMO.Word`; this package does not maintain a second implementation.

## Fidelity boundary

Common paragraphs, rich runs, tables, nested tables, images, notes, headers/footers, sections, styles, numbering, links, bookmarks, revisions, and comments map through the bridge. Unsupported Word elements and unsupported RTF objects/shapes are reported as omissions. Use `RequireNoLoss()` when those omissions must stop the workflow.

Ordinary uncropped inline RTF pictures in headers and footers retain the legacy VML picture carrier when written to Word. The carrier preserves their exact-line overhang through saving, reopening and subsequent PDF conversion. The destination's modern compatibility mode is unchanged. Body, note, cell, text-box and cropped-picture paths retain their existing DrawingML mapping.

Paragraph mapping retains minimum, exact, and multiple line spacing, including the sign and units described in the [RTF paragraph contract](../OfficeIMO.Rtf/README.md#paragraph-spacing-and-inheritance). The same effective-formatting path applies supported spacing, indentation, alignment, and tab stops to body paragraphs, table cells, notes, and running stories. Parsed RTF retains its active spacing and tab state; API-authored RTF paragraphs can inherit those properties from their styles. Missing non-auto before/after gaps in parsed paragraphs map to explicit zero, so the destination document's paragraph defaults do not add spacing. Explicit source values and automatic spacing remain in effect.

An RTF paragraph's active tab list replaces inherited tab lists. The bridge adds Word tab clears for unused positions declared anywhere in the associated paragraph-style chain, while retaining active positions and their alignment and leaders.

An explicitly selected default font from the RTF font table becomes the Word document's default font. Paragraph styles and individual runs can override it. An unset or unresolved font identifier retains the destination's existing default policy.

Table import retains a borderless, zero-padding source instead of adding Word table preset borders and margins. An explicit, uniform nonnegative row half-gap maps to direct left/right Word table cell margins. New Word documents use modern layout, where a uniform source row indent maps directly to the Word table indent. Legacy Word layout includes the half-gap in that stored indent; reverse conversion accounts for the source document's compatibility mode. Explicit cell padding overrides row padding and the half-gap, including zero.

Word-to-RTF recovers a half-gap from symmetric direct table margins in twips and a representable indent; other direct margins remain row padding. The qualified borderless left-aligned profile retains its content origin across modern mapping and legacy native import. Word ignores table indentation for centered and right-aligned tables; their placement follows the selected Word layout mode. These mappings do not establish native geometry for arbitrary padding, borders or alignments. Word spacing between separated cells is omitted with `WordRtfTableCellSpacingOmitted`. Missing, mixed, negative or unrepresentable RTF row half-gaps use zero table margins and report `RtfWordTableCellGapFlattened`. See the [support matrix](../Docs/officeimo.rtf-support-matrix.md#paragraph-layout-boundaries) for the supported ranges.

`NormalizeTablesForOnline()` preserves explicit host-cell top/bottom margins, including zero, and direct table/cell top-border declarations, including border suppression. It supplies fallbacks only where those values are absent.

Parsed legacy RTF can project ordinary note destinations to section-end Word endnotes. Modern note mode retains ordinary footnotes; `\ftnalt` marks endnotes. The [RTF note-mode policy](../OfficeIMO.Rtf/README.md) distinguishes the effective projection from the unchanged raw source note kind and settings.

Paragraph frames become ordinary text flow rather than positioned frames. RTF-to-Word reports `RtfWordParagraphFramesFlattened`; Word-to-RTF reports `WordRtfParagraphFramesFlattened`. Word tab-stop clearing instructions (`w:tab` with `w:val="clear"`) are omitted with `WordRtfTabClearsOmitted`; supported active tab stops are mapped separately. Word-side diagnostics include paragraph properties in document defaults, styles and individual stories. These are conversion losses, so `RequireNoLoss()` rejects them.

Legacy RTF paragraph numbering maps to editable Word lists for single-level decimal numbers and text bullets. Authored starts, cached numbering jumps, marker prefixes and suffixes, and basic marker formatting are retained. Body and table paragraphs share their numbering sequence; each header, footer, and note starts an independent sequence. Conversion leaves the source RTF model unchanged.

Legacy multilevel and continuation forms, special table numbering, spacing controls, and other numeric formats use cached or resolved marker text and report `RtfWordLegacyNumberingFlattened`. `RequireNoLoss()` rejects these conversions because editable numbering or marker formatting is not fully retained.

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages.
- **OfficeIMO:** `OfficeIMO.Word` and `OfficeIMO.Rtf`; Word workflows are reused instead of reimplemented.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 2 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Word.Rtf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
