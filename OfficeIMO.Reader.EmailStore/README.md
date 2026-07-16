# OfficeIMO.Reader.EmailStore

`OfficeIMO.Reader.EmailStore` adds PST, OST, OLM, and EMLX inputs to an isolated
`OfficeDocumentReader`. Parsing remains owned by `OfficeIMO.Email.Store`; this package only
registers the formats and projects parsed `EmailDocument` items through Reader's existing
email chunks, metadata, diagnostics, attachments, hashes, and rich result envelope.

## Install

```powershell
dotnet add package OfficeIMO.Reader.EmailStore
```

Install `OfficeIMO.Email.Store` directly instead when the application wants folder and item models without the Reader projection. Applications that want every local Reader adapter can install `OfficeIMO.Reader.All` and call `AddAllOfficeIMOHandlers()`.

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
Reader asset collection. The adapter opens an `EmailStoreSession`, selects at most 1,000 items by default, and only
then projects full messages into Reader. Configure parser limits, attachment retention, associated items,
or a PST password through `ReaderEmailStoreOptions.StoreOptions`. A Reader-level
`MaxInputBytes` can further narrow the registered store limit but cannot widen it.

For a huge PST or OST, query summaries before Reader reads bodies and attachments:

```csharp
var emailStore = new ReaderEmailStoreOptions {
    Query = new OfficeIMO.Email.Store.EmailStoreQuery(
        subjectContains: "invoice",
        since: DateTimeOffset.UtcNow.AddMonths(-6),
        maxItemsScanned: 500_000,
        maxResults: 250),
    MaxItems = 250,
    StoreOptions = new OfficeIMO.Email.Store.EmailStoreReaderOptions(
        retainAttachmentContent: false)
};

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddEmailStoreHandler(emailStore)
    .Build();
```

`SelectionLimitReached` metadata and `EMAIL_STORE_READER_SELECTION_LIMIT` diagnostics make bounded output visible.
An unreadable selected item is diagnosed and skipped by default; set `ContinueOnItemError = false` for fail-fast
ingestion.

Reader normally computes source and chunk hashes. This adapter leaves complete-store hashing disabled by default,
because hashing a huge PST/OST would force a full sequential read even for a selective query. Chunk hashes still
follow `ReaderOptions.ComputeHashes`. Set `ComputeSourceHash = true` only when the host explicitly wants that cost.

The implementation is fully managed and does not add a native or third-party email parser.

## Boundaries

- `OfficeIMO.Email.Store` owns PST/OST/OLM/EMLX parsing, mailbox-directory sessions, and the store model.
- `OfficeIMO.Reader.EmailStore` owns registration, stable logical paths, Reader limits, and result projection only.
- `OfficeIMO.Reader` continues to handle individual EML, MSG, OFT, TNEF, and mbox artifacts without this package.

Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.
