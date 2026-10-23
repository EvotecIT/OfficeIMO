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

Paragraph mapping retains minimum, exact, and multiple line spacing, including the sign and units described in the [RTF paragraph contract](../OfficeIMO.Rtf/README.md#paragraph-spacing-and-inheritance). The same effective-formatting path applies supported spacing, indentation, alignment, and tab stops to body paragraphs, table cells, notes, and running stories. Parsed RTF retains its active spacing and tab state; API-authored RTF paragraphs can inherit those properties from their styles.

An RTF paragraph's active tab list replaces inherited tab lists. The bridge adds Word tab clears for unused positions declared anywhere in the associated paragraph-style chain, while retaining active positions and their alignment and leaders.

An explicitly selected default font from the RTF font table becomes the Word document's default font. Paragraph styles and individual runs can override it. An unset or unresolved font identifier retains the destination's existing default policy.

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
