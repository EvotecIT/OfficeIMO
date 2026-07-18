# OfficeIMO.Reader.EmailStore

`OfficeIMO.Reader.EmailStore` is a thin PST, OST, OLM, and EMLX adapter package for
`OfficeIMO.Reader`. Parsing remains owned by the `OfficeIMO.Email.Store` API; this adapter only
registers the formats and projects parsed `EmailDocument` items through Reader's existing
email chunks, metadata, diagnostics, attachments, hashes, and rich result envelope.

## Install

```powershell
dotnet add package OfficeIMO.Reader.EmailStore
```

Install `OfficeIMO.Email` directly instead when the application wants folder and item models without the Reader projection. Applications that want every local Reader adapter can install `OfficeIMO.Reader.All` and call `AddAllOfficeIMOHandlers()`.

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

For ingestion pipelines, stream one selected store item at a time instead of returning one aggregate result:

```csharp
var options = new ReaderEmailStoreOptions {
    MaxItems = 10_000,
    StreamAttachmentContent = true,
    StoreOptions = new OfficeIMO.Email.Store.EmailStoreReaderOptions(
        retainAttachmentContent: false)
};

foreach (ReaderEmailStoreItemResult item in reader.ReadEmailStoreItems(
    "archive.pst",
    new ReaderOptions { ComputeHashes = false },
    options)) {
    foreach (ReaderChunk chunk in item.Chunks) {
        Console.WriteLine($"{chunk.Id}: {chunk.Text.Length} characters");
    }
}
```

The enumeration owns the store session and disposes it when enumeration ends, so consume any deferred attachment
content before moving that work outside the loop. HTML and RTF bodies use Reader's registered semantic handlers;
recognized attachments use the same modular extraction path as standalone files. `ItemDiagnostics` determines
`Succeeded`. Store-open and hierarchy diagnostics appear once on the first result as `StoreDiagnostics`, while
`Diagnostics` remains the combined view.

The implementation is fully managed and does not add a native or third-party email parser.

## Boundaries

- `OfficeIMO.Email.Store` owns PST/OST/OLM/EMLX parsing, mailbox-directory sessions, and the store model.
- `OfficeIMO.Reader.EmailStore` owns registration, stable logical paths, Reader limits, and result projection only.
- `OfficeIMO.Reader` handles individual EML, MSG, OFT, TNEF, and mbox artifacts without an explicit Store registration.

Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.

The adapter depends on `OfficeIMO.Reader` and the unified `OfficeIMO.Email` package. It adds no parser, native runtime,
Outlook automation, or third-party email-store dependency of its own.
