# OfficeIMO.Reader.Text (Preview)

`OfficeIMO.Reader.Text` is now a compatibility orchestrator for structured text adapters:
- delegates `.csv`/`.tsv` to `OfficeIMO.Reader.Csv`
- delegates `.json` to `OfficeIMO.Reader.Json`
- delegates `.xml` to `OfficeIMO.Reader.Xml`
- keeps a single legacy registration entry point for existing consumers

Registration into `OfficeIMO.Reader`:

```csharp
using OfficeIMO.Reader.Text;

DocumentReaderTextRegistrationExtensions.RegisterStructuredTextHandler(replaceExisting: true);
```

For new integrations, prefer dedicated handlers:
- `DocumentReaderCsvRegistrationExtensions.RegisterCsvHandler(...)`
- `DocumentReaderJsonRegistrationExtensions.RegisterJsonHandler(...)`
- `DocumentReaderXmlRegistrationExtensions.RegisterXmlHandler(...)`

Status:
- compatibility wrapper and intentionally non-packable/non-publishable
