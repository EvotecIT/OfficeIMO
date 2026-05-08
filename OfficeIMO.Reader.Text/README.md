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

Compatibility note:
- `RegisterStructuredTextHandler(...)` is intentionally manual-only and not included in registrar auto-discovery.

Status:
- packaged as `OfficeIMO.Reader.Text`
- compatibility wrapper for existing integrations; new integrations should prefer the dedicated CSV, JSON, and XML adapters
