# OfficeIMO.Reader.Html (Preview)

`OfficeIMO.Reader.Html` is a modular adapter for HTML ingestion.

Current scope:
- HTML -> Markdown (via `OfficeIMO.Markdown.Html`)
- Markdown chunk emission in `ReaderChunk` shape
- heading-aware chunk metadata (`Location.HeadingPath`, `Location.StartLine`) when `ReaderOptions.MarkdownChunkByHeadings = true`
- path and stream dispatch via `DocumentReader` handler registration
- `ReaderHtmlOptions.HtmlToMarkdownOptions` pass-through for markdown writer profiles, input limits, transforms, custom element converters, and visual round-trip hints
- `ReaderHtmlOptions.CreateOfficeIMOProfile()`, `CreatePortableProfile()`, and `CreateUntrustedHtmlProfile(maxInputCharacters)` helpers for reusable adapter profiles
- `ReaderHtmlOptions.Clone()` for safe option-template reuse during handler registration and direct reads
- warning chunk when HTML yields no markdown content

Registration into `OfficeIMO.Reader`:

```csharp
using OfficeIMO.Reader.Html;

DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler();
```

For untrusted or size-sensitive HTML, pass `ReaderHtmlOptions` during direct reads or handler registration:

```csharp
using OfficeIMO.Reader.Html;

DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler(
    htmlOptions: ReaderHtmlOptions.CreateUntrustedHtmlProfile(maxInputCharacters: 100_000),
    replaceExisting: true);
```

Use `ReaderHtmlOptions.CreatePortableProfile()` when the reader output should favor portable Markdown serialization. For custom ingestion contracts, set `HtmlToMarkdownOptions` directly or clone a profile and add transforms, element converters, or visual round-trip hints.

Status:
- packaged as `OfficeIMO.Reader.Html`
- preview-scoped modular adapter for `OfficeIMO.Reader`
