# OfficeIMO.Reader.Csv (Preview)

`OfficeIMO.Reader.Csv` is a modular CSV/TSV ingestion adapter for `OfficeIMO.Reader`:
- CSV/TSV chunking with table-aware output
- path and stream dispatch
- deterministic chunk IDs and row-based locations
- `MaxInputBytes` enforcement via shared `ReaderInputLimits`

Registration into `OfficeIMO.Reader`:

```csharp
using OfficeIMO.Reader.Csv;

DocumentReaderCsvRegistrationExtensions.RegisterCsvHandler(replaceExisting: true);
```

Status:
- scaffolded and intentionally non-packable/non-publishable
