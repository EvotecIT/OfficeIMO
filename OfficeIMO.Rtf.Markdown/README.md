# OfficeIMO.Rtf.Markdown

First-party semantic conversion between OfficeIMO RTF documents and Markdown documents.

```csharp
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;

var rtf = new RtfDocument();
rtf.AddParagraph("Hello **from RTF**");

string markdown = rtf.ToMarkdown();
RtfDocument roundTripped = markdown.ToRtfDocumentFromMarkdown();
```

This package deliberately converts document meaning rather than raw RTF control words. Use `OfficeIMO.Rtf` for syntax-level parsing, writing, diagnostics, and lossless RTF editing.
