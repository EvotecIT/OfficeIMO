# OfficeIMO.Reader.Html (Preview)

`OfficeIMO.Reader.Html` is a modular adapter for HTML ingestion.

Current scope:
- HTML -> Markdown (via `OfficeIMO.Markdown.Html`)
- Markdown chunk emission in `ReaderChunk` shape
- heading-aware chunk metadata (`Location.HeadingPath`, `Location.StartLine`) when `ReaderOptions.MarkdownChunkByHeadings = true`
- path and stream dispatch via `DocumentReader` handler registration
- warning chunk when HTML yields no markdown content

Registration into `OfficeIMO.Reader`:

```csharp
using OfficeIMO.Reader.Html;

DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler();
```

Status:
- scaffolded and intentionally non-packable/non-publishable
