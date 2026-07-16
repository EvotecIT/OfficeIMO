# OfficeIMO.Reader.EmailStore

`OfficeIMO.Reader.EmailStore` adds PST, OST, OLM, and EMLX inputs to an isolated
`OfficeDocumentReader`. Parsing remains owned by `OfficeIMO.Email.Store`; this package only
registers the formats and projects parsed `EmailDocument` items through Reader's existing
email chunks, metadata, diagnostics, attachments, hashes, and rich result envelope.

## Install

```powershell
dotnet add package OfficeIMO.Reader.EmailStore
```

Install `OfficeIMO.Email.Store` directly instead when the application wants folder and message models without the Reader projection. Applications that want every local Reader adapter can install `OfficeIMO.Reader.All` and call `AddAllOfficeIMOHandlers()`.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.EmailStore;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddEmailStoreHandler()
    .Build();

OfficeDocumentReadResult result = reader.ReadDocument("archive.pst");
```

Store folders become stable logical paths such as
`archive.pst!/Mailbox/Inbox/item-000012`. Attachments are exposed through the normal
Reader asset collection. Configure parser limits, attachment retention, associated items,
or a PST password through `ReaderEmailStoreOptions.StoreOptions`. A Reader-level
`MaxInputBytes` can further narrow the registered store limit but cannot widen it.

The implementation is fully managed and does not add a native or third-party email parser.

## Boundaries

- `OfficeIMO.Email.Store` owns PST/OST/OLM/EMLX parsing and the store model.
- `OfficeIMO.Reader.EmailStore` owns registration, stable logical paths, Reader limits, and result projection only.
- `OfficeIMO.Reader` continues to handle individual EML, MSG, OFT, TNEF, and mbox artifacts without this package.

Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.
