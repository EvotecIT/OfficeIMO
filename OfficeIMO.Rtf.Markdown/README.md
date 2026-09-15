# OfficeIMO.Rtf.Markdown

`OfficeIMO.Rtf.Markdown` provides semantic conversion between `RtfDocument` and `MarkdownDoc`.

```csharp
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;

RtfDocument rtf = RtfDocument.Load("input.rtf").Document;
var options = new RtfToMarkdownOptions {
    ImagePathFactory = (_, index) => $"media/image-{index + 1}.png",
    ImageExporter = (image, _, path) => {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, image.Data);
    }
};

RtfConversionResult<string> result = rtf.ToMarkdownResult(options);
string markdown = result.RequireNoLoss();
```

Footnotes and endnotes become Markdown footnote references and definitions. Tables, lists, rich inline formatting, links, and supported images have semantic mappings. Nested tables are flattened inside Markdown table cells; annotations and headers/footers are diagnostic omissions.

Convert the other direction with `markdown.ToRtfDocumentFromMarkdown()` or `MarkdownDoc.ToRtfDocument()`.

This bridge converts document meaning, not raw control words. Use `OfficeIMO.Rtf` lossless APIs when the original RTF syntax must remain exact.

## Dependency footprint

- **External:** None beyond the dependencies of its OfficeIMO format packages.
- **OfficeIMO:** `OfficeIMO.Rtf`, `OfficeIMO.Markdown`, and `OfficeIMO.Core` own parsing, semantic mapping, image export hooks, and loss reports.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 2 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Rtf.Markdown` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
