# OfficeIMO.Reader.Html (Preview)

`OfficeIMO.Reader.Html` is a modular adapter for HTML ingestion.

Current scope:
- HTML -> Word (via `OfficeIMO.Word.Html`)
- Word -> Markdown (via `OfficeIMO.Word.Markdown`)
- Markdown chunk emission in `ReaderChunk` shape

Registration into `OfficeIMO.Reader`:

```csharp
using OfficeIMO.Reader.Html;

DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler();
```

Status:
- scaffolded and intentionally non-packable/non-publishable
