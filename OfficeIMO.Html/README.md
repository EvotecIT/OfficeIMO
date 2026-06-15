# OfficeIMO.Html

`OfficeIMO.Html` contains shared HTML ingestion primitives used by OfficeIMO converters.

It owns the reusable parts that should behave consistently across HTML-to-Markdown, HTML-to-Word, and HTML-backed PDF workflows:

- URL policy evaluation and base URI resolution
- AngleSharp document parsing helpers
- image source discovery for `img`, lazy-loading attributes, `srcset`, and `picture/source`
- image data URI parsing and media-type extension mapping

It does not own output-specific rendering. Markdown AST creation, Word document generation, and PDF orchestration stay in their converter packages.

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
