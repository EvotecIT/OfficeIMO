# OfficeIMO.Word.Rtf

`OfficeIMO.Word.Rtf` maps RTF documents to and from `OfficeIMO.Word`. It also exposes RTF-facing document workflows that reuse the Word engines and return a combined fidelity report.

## Convert with diagnostics

```csharp
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;

RtfDocument rtf = RtfDocument.Load("input.rtf").Document;
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
