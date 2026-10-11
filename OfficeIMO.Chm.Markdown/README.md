# OfficeIMO.Chm.Markdown

Recover a compiled help book as linked Markdown through `OfficeIMO.Chm` and the canonical `OfficeIMO.Markdown.Html` converter.

```csharp
using OfficeIMO.Chm;

ChmDocument book = ChmDocument.Load("manual.chm");
ChmConversionResult<string> result = book.ToMarkdownResult();
File.WriteAllText("manual.md", result.RequireValue());
foreach (var diagnostic in result.Report.FidelityDiagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
```

Use `ToMarkdownDocumentResult()` for an editable `MarkdownDoc` from the canonical Markdown engine. Pass `ChmConversionOptions` to select topic paths and bound aggregate HTML, images and serialized output. Pass `HtmlToMarkdownOptions` to choose the existing Markdown conversion/write policy. `ToMarkdown()` is a convenience wrapper; use the result API to retain fidelity evidence or call `Report.RequireNoLoss()` for strict acceptance.

Topic headings and links receive unique book anchors. Archive images are embedded as data URLs by default. Contents order is retained; compiled keyword-index and help-viewer features remain exposed by the source CHM model. Layout, styles and unsupported HTML structures follow the Markdown engine's reported conversion boundaries. See [CHM support](../OfficeIMO.Chm/SUPPORT.md).

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Chm.Markdown` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
