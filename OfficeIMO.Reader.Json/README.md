# OfficeIMO.Reader.Json (Preview)

`OfficeIMO.Reader.Json` is a modular JSON ingestion adapter for `OfficeIMO.Reader`:
- AST traversal (`System.Text.Json`) to path/type/value rows
- chunked structured output with optional markdown tables
- path and stream dispatch
- warning chunks for malformed JSON

Registration into `OfficeIMO.Reader`:

```csharp
using OfficeIMO.Reader.Json;

DocumentReaderJsonRegistrationExtensions.RegisterJsonHandler(replaceExisting: true);
```

Status:
- scaffolded and intentionally non-packable/non-publishable
