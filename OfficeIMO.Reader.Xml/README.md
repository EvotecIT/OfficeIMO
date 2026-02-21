# OfficeIMO.Reader.Xml (Preview)

`OfficeIMO.Reader.Xml` is a modular XML ingestion adapter for `OfficeIMO.Reader`:
- XML tree traversal to element/attribute path rows
- chunked structured output with optional markdown tables
- path and stream dispatch
- warning chunks for malformed XML

Registration into `OfficeIMO.Reader`:

```csharp
using OfficeIMO.Reader.Xml;

DocumentReaderXmlRegistrationExtensions.RegisterXmlHandler(replaceExisting: true);
```

Status:
- scaffolded and intentionally non-packable/non-publishable
