# OfficeIMO.Rtf.Html

Dependency-free semantic conversion between HTML and the OfficeIMO RTF document model.

```csharp
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Html;

RtfDocument document = "<p>Hello <strong>RTF</strong></p>".ToRtfDocumentFromHtml();
string html = document.ToHtml();
```

RTF-to-RTF editing in `OfficeIMO.Rtf` is the lossless preservation path. HTML conversion is a semantic bridge: it preserves supported text, inline formatting, links, lists, tables, bookmarks, fields, form fields, notes, tracked revisions, object metadata, shape metadata, and embedded PNG/JPEG images without pulling in external parsers or Office/COM automation.
