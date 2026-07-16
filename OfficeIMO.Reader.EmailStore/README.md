# OfficeIMO.Reader.EmailStore

`OfficeIMO.Reader.EmailStore` adds PST, OST, OLM, and EMLX inputs to an isolated
`OfficeDocumentReader`. Parsing remains owned by `OfficeIMO.Email.Store`; this package only
registers the formats and projects parsed `EmailDocument` items through Reader's existing
email chunks, metadata, diagnostics, attachments, hashes, and rich result envelope.

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
