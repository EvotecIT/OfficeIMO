# OfficeIMO.Reader.Epub (Preview)

`OfficeIMO.Reader.Epub` bridges `OfficeIMO.Epub` output into `OfficeIMO.Reader` chunk contracts.

Current scope:
- chapter-to-chunk projection
- max-char chunk splitting
- markdown + text chunk payloads
- warning chunks propagated from EPUB parser warnings
- virtual source paths (`.epub::chapter.xhtml`) for traceability
- path and stream dispatch via `DocumentReader` handler registration

Registration into `OfficeIMO.Reader`:

```csharp
using OfficeIMO.Reader.Epub;

DocumentReaderEpubRegistrationExtensions.RegisterEpubHandler();
```

Status:
- scaffolded and intentionally non-packable/non-publishable
