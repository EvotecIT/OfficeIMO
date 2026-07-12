# OfficeIMO breaking API migration

This release removes compatibility aliases and standardizes persistence, conversion, diagnostics, themes, and asynchronous I/O across the OfficeIMO packages. The cleanup is intentionally breaking: migrate to the canonical names instead of adding local shims.

## Document persistence

Word, Excel, PowerPoint, Visio, RTF, PDF, and OpenDocument models now follow the same basic vocabulary:

| Intent | API |
| --- | --- |
| Save to the document's current path | `Save()` / `SaveAsync()` |
| Save to a path or stream | `Save(pathOrStream)` / `SaveAsync(pathOrStream)` |
| Produce package bytes | `ToBytes()` |
| Produce an in-memory stream | `ToStream()` or the format-specific canonical stream method |
| Export to another format | `ToPdf()`, `ToHtml()`, `ToMarkdown()`, and matching `SaveAs{Format}()` methods |

The shared `DocumentAccessMode`, `DocumentPersistenceMode`, `DocumentCreateOptions`, and `DocumentLoadOptions`
contracts are owned by the existing zero-dependency `OfficeIMO.Drawing` foundation. Import `OfficeIMO.Drawing` when
configuring Word, Excel, or PowerPoint lifecycle options. There is no separate `OfficeIMO.Core` package.

`SaveAsync` exists only where the destination performs asynchronous I/O. Pure in-memory conversion and byte generation are synchronous. Use `ToBytes()` or `ToPdf()` directly; do not wrap them in a removed `*Async` compatibility method.

Streams used as associated destinations must be writable and seekable so every parameterless `Save()` can replace the complete artifact. Editable `Load(Stream)` calls retain only streams with those capabilities; read-only and non-seekable inputs remain detached and require an explicit destination. `Save(Stream)` and `SaveAsync(Stream)` are one-time writes and do not silently redirect later parameterless saves. Use `Create(Stream)` or load an editable seekable stream when persistent stream association is intended.

OpenDocument callers that need save diagnostics should use `SaveResult()`, `SaveResultAsync()`, `ToBytesResult()`, or `SaveFlatXmlResult()`. The returned `OdfSaveResult` exposes `Value`, `Report`, `HasLoss`, `RequireValue()`, and `RequireNoLoss()`.

```csharp
OdfSaveResult result = document.SaveResult("output.odt");
result.RequireNoLoss();

foreach (string entry in result.Report.RewrittenEntries) {
    Console.WriteLine(entry);
}
```

## Conversion results and diagnostics

Reusable options are configuration only. They no longer retain reports or warning lists from the last operation. Request the result-bearing method when diagnostics matter:

```csharp
PdfDocumentConversionResult conversion = document.ToPdfDocumentResult(options);
PdfSaveResult save = conversion.TrySave("output.pdf");

foreach (PdfConversionWarning warning in save.Warnings) {
    Console.WriteLine($"{warning.Code}: {warning.Message}");
}
```

Result objects consistently expose `Value` and `Report`. PDF save attempts expose `Report`, `Warnings`, and `HasWarnings` in addition to write success/failure. RTF bridges use `RtfConversionResult<T>` with `Value`, `Report`, `RequireValue()`, and `RequireNoLoss()`.

The canonical PDF conversion-result method is `ToPdfDocumentResult()`. Source-explicit overloads use names such as `ToPdfDocumentFromMarkdownResult()`, `ToPdfDocumentFromRtfResult()`, and `ToWordDocumentFromPdfResult()`.

## Renamed and removed members

| Removed member | Replacement |
| --- | --- |
| `WordImage.SaveToFile(...)` | `WordImage.Save(...)` |
| `WordImage.GetBytes()` | `WordImage.ToBytes()` |
| `WordImage.GetStream()` | `WordImage.OpenRead()` |
| `WordDocument.GetImages()` / `GetImageStreams()` | `GetImageBytes()` / `OpenImageStreams()` |
| `ExcelImage.GetBytes()` | `ExcelImage.ToBytes()` |
| `WordComment.Delete()` | `WordComment.Remove()` |
| `WordTable.AutoFit` | `WordTable.LayoutMode` |
| `ExcelDocument.MergeWorkSheets(...)` / `JoinWorkSheets(...)` | `ExcelDocument.MergeWorksheets(...)` |
| `ExcelDocument.CompareWorkSheets(...)` | `ExcelDocument.CompareWorksheets(...)` |
| `ExcelDocument.JoinWorkbookFrom(...)` | `ExcelDocument.MergeWorksheets(...)` |
| `ExcelDocument.CreateTableOfContents(...)` | `ExcelDocument.AddTableOfContents(...)` |
| `ExcelSheet.SetCellValues(...)` | `ExcelSheet.CellValues(...)` |
| `ExcelSheet.CellValuesParallel(...)` | `ExcelSheet.CellValues(..., ExecutionMode.Parallel)` |
| `VisioDocument.UseMastersFromTemplate(...)` | `VisioDocument.LearnMastersFromVsdx(...)` |
| `ListItem.BlockChildren` | `ListItem.ChildBlocks` |
| `MarkdownDoc.SaveHtml(...)` | `MarkdownDoc.SaveAsHtml(...)` |
| `OutlookContact.Email1Address` | `OutlookContact.Email1.Address` |
| `OutlookContact.BusinessPhone` and related phone aliases | `OutlookContact.Phones` |
| `TrackComments` | no replacement; it incorrectly toggled revision tracking. Use `TrackChanges` or `Settings.TrackRevisions` when revision tracking is intended |
| `ToPdfResult()` | `ToPdfDocumentResult()` |
| `ToWordResult()` for PDF input | `ToWordDocumentFromPdfResult()` |
| `PdfSaveResult.ConversionWarnings` / `PdfBytesResult.ConversionWarnings` | `Warnings` and `Report` |
| `RtfDocument.ToMemoryStream()` | `RtfDocument.ToStream()` |
| `ToRtfMemoryStream()` / `ToRtfMemoryStreamAsync()` | `ToRtfStream()` / `ToRtfStreamAsync()` |
| `SavePdfAsWord()` / `SavePdfAsRtf()` | `SaveAsWordFromPdf()` / `SaveAsRtfFromPdf()`; use the `*FromPdfFile()` form for a source path |

Compatibility-only members such as `LastSaveReport`, public `LastSaveDiagnostics`, option-owned `ConversionReport`, and option-owned `Warnings` were removed. Use the operation result instead.

`ExcelSaveDiagnostics` and `ExcelSavePackageWriter` are internal implementation details. They had no public operation-result producer after `LastSaveDiagnostics` was removed, so retaining them as public types created an unusable contract.

The generic `Helpers` file-copy and `IsFileLocked` methods were also removed. Use `File.ReadAllBytes`, `File.OpenRead`, `File.Copy`, and `Stream.CopyTo` directly; filesystem lock probing belongs in application or test code. These wrappers added no Office document behavior.

Color-to-hex formatting is owned by `OfficeIMO.Drawing`. Import that namespace and use its `ToHexColor()` extension (or `OfficeColor.ToRgbHex()`); the duplicate Word and Excel helper extensions were removed.

Hexadecimal Office color values are normalized to uppercase `RRGGBB` without a leading `#`. Word no longer applies its former package-local lowercase convention, and legacy `.doc` palette conversion accepts the canonical representation.

## Theme naming

Markdown uses one shared cross-format `MarkdownVisualTheme` through `Theme`:

```csharp
var htmlOptions = new HtmlOptions { Theme = MarkdownVisualTheme.Report() };
var pdfOptions = new MarkdownPdfSaveOptions { Theme = MarkdownVisualTheme.Report() };
```

PDF-only visual details use `MarkdownPdfSaveOptions.PdfTheme`. The canonical helpers are `ApplyDefaultTheme()` and `UseFrontMatterTheme`. The removed names `VisualTheme`, `ApplyWordLikeTheme()`, and `UseFrontMatterVisualTheme` should not be preserved in consumer wrappers.

`HtmlOptions` is reusable configuration. Rendering and save operations clone its nested settings, so `ToHtmlFragment()`, `ToHtmlDocument()`, `ToHtmlParts()`, and concurrent HTML saves no longer change the caller's output kind or retain operation state. `SaveAsHtmlAsync(...)` accepts a cancellation token and performs asynchronous file I/O.

## Async migration

Removed async methods performed no asynchronous work. Replace them with their synchronous counterparts:

| Removed async member | Replacement |
| --- | --- |
| `MarkdownDoc.ToHtmlFragmentAsync()` | `ToHtmlFragment()` |
| `MarkdownDoc.ToHtmlDocumentAsync()` | `ToHtmlDocument()` |
| `WordDocument.ToMarkdownAsync()` | `ToMarkdown()` |
| `WordDocument.ToMarkdownDocumentAsync()` | `ToMarkdownDocument()` |
| `MarkdownDoc.ToWordDocumentAsync()` | `ToWordDocument()` |
| in-memory HTML/RTF `ToHtmlAsync()`, `ToRtfAsync()`, and byte/stream variants | the matching synchronous conversion method |
| `RtfDocument.ReadAsync(string)` | `RtfDocument.Read(string)` |
| `RtfDocument.LoadAsync(byte[])` | `RtfDocument.Load(byte[])` |
| RTF `ToRtfAsync()`, `ToBytesAsync()`, and lossless byte async methods | the matching synchronous method |
| Word/RTF or Word/PDF byte-returning async conversion | `ToRtf()`, `ToRtfBytes()`, or `ToPdf()` |

Async file and stream reads/writes remain available, including `LoadAsync(pathOrStream)`, `SaveAsync(pathOrStream)`, and `SaveAs{Format}Async(pathOrStream)`.

## Reader and adapter ownership

`OfficeIMO.OpenDocument` owns native ODF package behavior. Word, Excel, and PowerPoint OpenDocument packages are thin projections over that owner. Reader options are reusable configuration; RTF read diagnostics are returned by `ReadRtfFileResult()`, `ReadRtfResult()`, `ReadRtfChunksResult()`, or the rich `OfficeDocumentReadResult` rather than being stored on `ReaderRtfOptions`.
