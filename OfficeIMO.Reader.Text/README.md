# OfficeIMO.Reader.Text (Preview)

`OfficeIMO.Reader.Text` is a modular path for improving non-Office text ingestion:
- CSV semantic chunking (table-aware)
- JSON AST chunking (path/type/value tables)
- XML AST chunking (element/attribute path tables)
- future structured text adapters

Registration into `OfficeIMO.Reader`:

```csharp
using OfficeIMO.Reader.Text;

DocumentReaderTextRegistrationExtensions.RegisterStructuredTextHandler(replaceExisting: true);
```

Status:
- scaffolded and intentionally non-packable/non-publishable
