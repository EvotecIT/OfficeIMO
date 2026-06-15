# OfficeIMO.Rtf.Html

Dependency-free semantic conversion between HTML and the OfficeIMO RTF document model.

```csharp
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Html;

RtfDocument document = "<p>Hello <strong>RTF</strong></p>".ToRtfDocumentFromHtml();
string html = document.ToHtml();
```

RTF-to-RTF editing in `OfficeIMO.Rtf` is the lossless preservation path. HTML conversion is a semantic bridge: it preserves supported text, inline formatting, links, lists, tables, bookmarks, fields, form fields, notes, tracked revisions, object metadata, shape metadata, and embedded PNG/JPEG images without pulling in external parsers or Office/COM automation.

For workflow systems that use HTML as an interchange surface, including clinical and document-review systems, OfficeIMO keeps RTF-only state in `officeimo-rtf-*` metadata and `data-officeimo-rtf-*` attributes. This gives the bridge a stable place to grow without pretending plain HTML can represent every RTF control word by itself.

Use `RtfHtmlReadOptions.CreateUntrustedHtmlProfile()` when importing HTML from bounded ingestion surfaces. The profile keeps conversion offline and applies node/depth limits, while `RtfHtmlReadOptions.Diagnostics` and `DiagnosticHandler` provide a stable place for skipped/degraded-content reporting as more healthcare and workflow-specific cases are added.
