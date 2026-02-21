# OfficeIMO.Reader.Zip (Preview)

`OfficeIMO.Reader.Zip` is a modular adapter that traverses ZIP entries and forwards supported entry types into `OfficeIMO.Reader`.

Current scope:
- safe entry enumeration via `OfficeIMO.Zip`
- best-effort entry extraction into `ReaderChunk`
- warning chunks for skipped/error entries
- bounded nested ZIP traversal with reusable `ReaderZipOptions`
- path and stream dispatch via `DocumentReader` handler registration
- non-seekable stream support (adapter snapshots input when required)

Registration into `OfficeIMO.Reader`:

```csharp
using OfficeIMO.Reader.Zip;

DocumentReaderZipRegistrationExtensions.RegisterZipHandler();
```

Status:
- scaffolded and intentionally non-packable/non-publishable
