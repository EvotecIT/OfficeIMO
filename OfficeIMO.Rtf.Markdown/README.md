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

string markdown = rtf.ToMarkdown(options);
options.ConversionReport.RequireNoLoss();
```

Footnotes and endnotes become Markdown footnote references and definitions. Tables, lists, rich inline formatting, links, and supported images have semantic mappings. Nested tables are flattened inside Markdown table cells; annotations and headers/footers are diagnostic omissions.

Convert the other direction with `markdown.ToRtfDocumentFromMarkdown()` or `MarkdownDoc.ToRtfDocument()`.

This bridge converts document meaning, not raw control words. Use `OfficeIMO.Rtf` lossless APIs when the original RTF syntax must remain exact.
