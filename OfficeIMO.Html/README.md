# OfficeIMO.Html

`OfficeIMO.Html` contains shared HTML ingestion primitives and first-party HTML bridge APIs used by OfficeIMO converters.

It owns the reusable parts that should behave consistently across HTML-to-Markdown, HTML-to-Word, HTML-to/from-RTF, and HTML-backed PDF workflows:

- URL policy evaluation and base URI resolution
- AngleSharp document parsing helpers
- DOM traversal facts and node/depth limit tracking
- image source discovery for `img`, lazy-loading attributes, `srcset`, and `picture/source`
- image data URI parsing and media-type extension mapping
- semantic HTML to/from RTF conversion over the dependency-free `OfficeIMO.Rtf` model

It does not replace output-specific engines. Markdown AST creation, Word document generation, RTF document generation, and PDF orchestration stay in their owning packages such as `OfficeIMO.Markdown.Html`, `OfficeIMO.Word.Html`, `OfficeIMO.Rtf`, and `OfficeIMO.Html.Pdf`.

## RTF Bridge

```csharp
using OfficeIMO.Html;

RtfDocument document = "<p>Hello <strong>RTF</strong></p>".ToRtfDocument();
string rtf = document.ToRtf();
string html = document.ToHtml();
```

RTF-to-RTF editing in `OfficeIMO.Rtf` remains the lossless preservation path. The HTML bridge is semantic: it preserves supported text, inline formatting, links, lists, tables, bookmarks, fields, form fields, notes, tracked revisions, object metadata, shape metadata, and embedded PNG/JPEG images without Office/COM automation.

## URL Policy

```csharp
var policy = HtmlUrlPolicy.CreateWebOnlyProfile();
string href = HtmlUrlPolicyEvaluator.ResolveUrl(
    "/docs/start.html",
    new Uri("https://example.com/"),
    policy);
```

## Parsing And Base URIs

```csharp
var document = HtmlDocumentParser.ParseDocument(html);
Uri? baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(
    document,
    new Uri("https://example.com/articles/"));
```

## Traversal Limits

```csharp
HtmlDomLimitTracker? tracker = HtmlDomLimitTracker.Create(
    maxHtmlNodes: 10000,
    maxHtmlDepth: 64);
```

Converter packages use these primitives to keep bounded HTML ingestion behavior consistent while still reporting converter-specific diagnostics.

## Shared Diagnostics And Gallery Contracts

```csharp
var report = new HtmlDiagnosticReport();
report.Add("OfficeIMO.Word.Html", "HtmlCommentSkipped", "Comment skipped");

var scenario = new HtmlCapabilityGalleryScenario(
    "quarterly-report",
    "Quarterly Report",
    "Word HTML",
    "HTML import, DOCX validation, and round-trip export proof");
```

`HtmlDiagnosticReport` and the capability-gallery contracts provide a common shape for HTML converters, PDF bridges, readers, tests, and future documentation generators. Format-specific packages can keep their existing compatibility APIs while also publishing shared diagnostics and artifact metadata for market-facing proof galleries.

## Image Sources

```csharp
string source = HtmlImageSourceResolver.ResolveImageSource(
    imageElement,
    baseUri,
    HtmlUrlPolicy.CreateOfficeIMOProfile());
```

## Image Data URIs

```csharp
if (HtmlImageDataUri.TryParse(source, out var dataUri) && dataUri.IsBase64) {
    byte[] bytes = dataUri.DecodeBytes();
    string extension = dataUri.FileExtension;
}
```
