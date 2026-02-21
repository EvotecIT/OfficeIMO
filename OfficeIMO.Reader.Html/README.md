# OfficeIMO.Reader.Html (Preview)

`OfficeIMO.Reader.Html` is a modular adapter for HTML ingestion.

Current scope:
- HTML -> Word (via `OfficeIMO.Word.Html`)
- Word -> Markdown (via `OfficeIMO.Word.Markdown`)
- Markdown chunk emission in `ReaderChunk` shape
- path and stream dispatch via `DocumentReader` handler registration
- warning chunk when HTML yields no markdown content

Registration into `OfficeIMO.Reader`:

```csharp
using OfficeIMO.Reader.Html;

DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler();
```

With custom conversion options:

```csharp
using OfficeIMO.Reader.Html;
using OfficeIMO.Word.Markdown;

DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler(
    htmlOptions: new ReaderHtmlOptions {
        MarkdownOptions = new WordToMarkdownOptions {
            EnableUnderline = true
        }
    },
    replaceExisting: true);
```

Status:
- scaffolded and intentionally non-packable/non-publishable
