# OfficeIMO.Email.Store

`OfficeIMO.Email.Store` is the fully managed mailbox-store package for PST, OST, OLM, EMLX, Apple Mail, and Maildir sources. Source stores remain read-only. Selected items project into the common `OfficeIMO.Email.EmailDocument` model, and applications can create a new Unicode PST or convert a supported source into a separate PST without Outlook, native libraries, or third-party parser packages.

## Install

```powershell
dotnet add package OfficeIMO.Email.Store
```

## Read a store

Open a session when the store may be large. Folder discovery and item enumeration do not materialize
every PST/OST item or attachment:

```csharp
using OfficeIMO.Email.Store;

using EmailStoreSession session = EmailStoreSession.Open("archive.pst");

EmailStoreFolderInfo inbox = session.Folders.Single(folder => folder.Name == "Inbox");
foreach (EmailStoreItemReference reference in session.EnumerateItems(
    new EmailStoreEnumerationOptions(folderId: inbox.Id, maxItems: 100))) {
    EmailStoreItem item = session.ReadItem(reference);
    Console.WriteLine(item.Document.Subject);
}
```

The session keeps a bounded B-tree page cache, streams NBT entries and large table row-matrix blocks,
and resolves individual NIDs and BIDs on demand. Sessions are not thread-safe. A caller-owned stream
is left open by default and its original position is restored when the session is disposed.

## Search a huge PST or OST

PST/OST search reads only a small summary-property allowlist until the application explicitly selects an item.
Both scanning and returned results are bounded:

```csharp
using EmailStoreSession session = EmailStoreSession.Open("mailbox.ost");

foreach (EmailStoreSearchResult match in session.Search(new EmailStoreQuery(
    folderId: session.Folders.Single(folder => folder.Name == "Inbox").Id,
    subjectContains: "quarterly report",
    since: DateTimeOffset.UtcNow.AddYears(-1),
    maxItemsScanned: 250_000,
    maxResults: 100))) {
    Console.WriteLine($"{match.Summary.ReceivedAt:u} {match.Summary.Subject}");

    EmailStoreItem item = session.ReadItem(match.Reference);
    // Body, recipients, and attachments are projected only here.
}
```

Opening and enumerating a PST/OST does not load the complete NBT, BBT, folder contents, message bodies, or
attachments. The default source limit is 1 TiB and can be changed explicitly. The default cache retains at most
512 B-tree pages. One selected item is still subject to per-item, property, and attachment limits; use
`retainAttachmentContent: false` for metadata/text ingestion so projected items do not retain attachment payloads.

Body, recipient, and attachment-name search is also bounded and resumable. It returns the matched fields, a
plain-text snippet, progress counts, diagnostics for skipped items, and a checkpoint for the next batch:

```csharp
EmailStoreContentSearchCheckpoint? checkpoint = null;
do {
    EmailStoreContentSearchReport batch = session.SearchContent(
        new EmailStoreContentQuery(
            new[] { "invoice", "contoso" },
            fields: EmailStoreContentSearchFields.Bodies |
                    EmailStoreContentSearchFields.AttachmentNames,
            matchMode: EmailStoreContentMatchMode.AllTerms,
            maxItemsScanned: 5_000,
            maxResults: 100,
            maxDecodedPropertyBytesPerItem: 16L * 1024 * 1024,
            maxSearchableCharactersPerItem: 500_000,
            resumeFrom: checkpoint));

    foreach (EmailStoreContentSearchResult match in batch.Results) {
        Console.WriteLine($"{match.Summary.Subject}: {match.Snippet}");
    }
    checkpoint = batch.NextCheckpoint;
} while (checkpoint != null);
```

Checkpoints are offsets in the chosen enumeration, so keep the same metadata filter and inclusion options between
batches. Reusing the same open session also reuses its bounded PST/OST page cache.

## Select item parts and stream attachments

`ReadItem` can load metadata, bodies, recipients, attachment metadata, attachment content, embedded items, or
extended MAPI properties independently. Attachment content can remain behind a reopenable, forward-only stream:

```csharp
var storeOptions = new EmailStoreReaderOptions(retainAttachmentContent: false);
using EmailStoreSession session = EmailStoreSession.Open("mailbox.ost", storeOptions);

EmailStoreItemReference reference = session.EnumerateItems(
    new EmailStoreEnumerationOptions(maxItems: 1)).Single();
EmailStoreItem item = session.ReadItem(reference, new EmailStoreItemReadOptions(
    EmailStoreItemReadParts.Metadata |
    EmailStoreItemReadParts.Bodies |
    EmailStoreItemReadParts.AttachmentMetadata |
    EmailStoreItemReadParts.AttachmentContent,
    maxDecodedPropertyBytes: 16L * 1024 * 1024,
    preferStreamingAttachmentContent: true));

foreach (OfficeIMO.Email.EmailAttachment attachment in item.Document.Attachments) {
    using Stream content = attachment.OpenContentStream();
    // Copy or inspect the payload while the EmailStoreSession remains open.
}
```

Opening a deferred attachment stream does not decode its payload. Bytes are pulled as the caller reads. The stream
becomes invalid when its owning session is disposed.

Use `EmailStoreReader` when the application explicitly wants the complete configured store scope in memory:

```csharp
using OfficeIMO.Email.Store;

EmailStoreReadResult result = new EmailStoreReader().Read("archive.pst");

foreach (EmailStoreFolder folder in result.Store.Folders) {
    foreach (EmailStoreItem item in folder.Items) {
        Console.WriteLine($"{folder.Name}: {item.Document.Subject}");
    }
}

foreach (EmailStoreDiagnostic diagnostic in result.Diagnostics) {
    Console.WriteLine($"{diagnostic.Severity}: {diagnostic.Code}: {diagnostic.Message}");
}
```

The same reader detects the format from a bounded header plus the source name:

```csharp
EmailStoreFormat format = EmailStoreReader.DetectFormat("export.olm");
EmailStoreReadResult result = new EmailStoreReader().Read("export.olm");
```

## Supported inputs

| Format | Current contract |
| --- | --- |
| PST | ANSI and Unicode NDB stores, folders, contents tables, ordinary and associated items, named properties, attachments, and embedded messages. |
| OST | The supported PST-compatible NDB paths plus compressed blocks used by supported OST variants. Server-only or unmaterialized content cannot be recovered from an offline file. |
| OLM | Bounded Outlook for Mac ZIP/XML archives, folders, messages and typed items, and safe in-archive attachments. |
| EMLX | One Apple Mail EMLX item, including its RFC 5322/MIME message and supported XML property-list metadata. Partial files report that external Apple Mail content may be absent. |
| Mailbox directory | Lazy Apple Mail `.mbox/.../Messages/*.emlx`, Maildir `cur/new`, and EML/MIME directory trees. Reparse points are not followed. |

PST and OST MAPI properties use the same projections as MSG and OFT, so messages, appointments, contacts, tasks, journals, notes, recipients, attachments, and named properties do not acquire a second public model.

Folder metadata includes `SpecialFolderKind`, the `ClassificationSource` used to establish it, the MAPI
`ContainerClass`, and `IsSearchFolder`. Provider identifiers take precedence; language-dependent display-name
matching is an explicit fallback. Selected items expose `ContentAvailability`, separating requested parts that are
available, unavailable, or indeterminate. This matters for OST headers and other content that was never cached locally, and
for partial EMLX artifacts whose sibling content is absent.

PST/OST and mailbox-directory sessions are lazy. Single EMLX input contains one item. OLM currently validates and
materializes its bounded ZIP/XML archive when the session opens; query, validation, and export then use the same
session contracts.

## Inspect, validate, and recover

Inspection reads the already-built catalog. Validation depth is explicit, and orphan discovery never mutates or
"repairs" the source:

```csharp
EmailStoreInspectionReport inspection = session.Inspect();

EmailStoreValidationReport validation = session.Validate(
    new EmailStoreValidationOptions(
        mode: EmailStoreValidationMode.Summaries,
        maxItems: 100_000,
        verifyStructuralIntegrity: true,
        maxStructuralPages: 100_000,
        maxStructuralBlocks: 100_000,
        maxStructuralBytes: 1024L * 1024 * 1024));

EmailStoreRecoveryReport recovery = session.DiscoverRecoverableItems(
    new EmailStoreRecoveryOptions(
        maxItemsScanned: 1_000_000,
        maxRecoveredItems: 10_000));
```

`Shallow` validation covers the header, indexes, and folder catalog. `Summaries` selectively decodes browsing
properties. `FullItems` projects bodies, recipients, and attachment metadata/content according to reader options.
A configured limit produces an incomplete report, not a false corruption claim. Opt-in PST/OST structural
validation checks BBT/NBT page layout and ordering, page and block bounds, BIDs, signatures, CRCs, and 4K OST stored
and decoded lengths. Page, block, and byte ceilings are reported separately, including whether validation stopped
at a bound.

## Export and migration

Directory export writes one selected item at a time as EML, MSG, OFT, or TNEF and records a tab-separated
preservation manifest. Streaming mbox export writes a same-directory temporary file and commits it only after the
selected sequence completes:

```csharp
EmailStoreExportReport files = session.ExportToDirectory(
    "exported-mail",
    new EmailStoreExportOptions(
        format: OfficeIMO.Email.EmailFileFormat.OutlookMsg,
        maxItems: 50_000));

EmailStoreMboxExportReport mailbox = session.ExportToMbox(
    "exported-mail/archive.mbox",
    new EmailStoreMboxExportOptions(maxItems: 50_000));
```

Output conversion uses `OfficeIMO.Email` and its explicit semantic-loss policy. Existing destinations are not
replaced by default. Per-item failures and fidelity warnings remain visible in the export report.

## Create a Unicode PST

`EmailStorePstWriter` creates a new PST incrementally and commits it through a same-directory temporary file. It
does not append to or edit an existing PST:

```csharp
using OfficeIMO.Email;
using OfficeIMO.Email.Store;

using EmailStorePstWriter writer = EmailStorePstWriter.Create(
    "created.pst",
    new EmailStorePstWriterOptions(displayName: "Project archive"));

string inbox = writer.AddFolder("Inbox", containerClass: "IPF.Note");
writer.AddItem(inbox, new EmailDocument {
    Subject = "Created without Outlook",
    MessageClass = "IPM.Note"
});

EmailStorePstWriteReport report = writer.Complete();
```

The writer owns Unicode NDB allocation and indexes, Heap-on-Node property and table contexts, folders, ordinary
and associated items, recipients, attachments, embedded messages, named properties, and fixed or variable
multi-valued MAPI properties. It reuses `OfficeIMO.Email` item projection so appointments, contacts, tasks,
journals, notes, and messages do not acquire a PST-only property model.

Set `failOnDataLoss: true` when a warning should prevent the final commit. Examples include an attachment whose
payload was not retained, structured-storage metadata without its original compound payload, or a named property
whose source Name-to-ID mapping was unavailable. `Diagnostics` identifies the affected item or property without
including message or attachment content.

## Convert OST or another store to a new PST

Conversion opens the source read-only and always writes a different destination:

```csharp
EmailStorePstConversionReport report = EmailStoreConverter.ConvertToPst(
    "mailbox.ost",
    "mailbox-converted.pst",
    conversionOptions: new EmailStorePstConversionOptions(
        includeAssociatedItems: true,
        includeOrphanedItems: true,
        failOnDataLoss: false));
```

The same API accepts every format supported by `EmailStoreSession`. Search-folder results can be copied as static
folders, but their dynamic query definitions are not regenerated. Content that was never cached in an OST cannot
be recovered from the offline file. The conversion report separates converted and skipped items and combines
reader, conversion, and writer diagnostics.

## Limits and attachment retention

Reads are bounded before large structures are retained. Applications can narrow source bytes, nodes, folders, items, properties, attachments, archive entries, XML characters, and embedded-message depth:

```csharp
var options = new EmailStoreReaderOptions(
    maxInputBytes: 2L * 1024 * 1024 * 1024,
    maxItemCount: 100_000,
    maxCachedBTreePages: 256,
    retainAttachmentContent: false,
    includeAssociatedItems: false);

EmailStoreReadResult result = new EmailStoreReader(options).Read("archive.ost");
```

For the materializing `EmailStoreReader`, `retainAttachmentContent: false` omits attachment payloads so the returned
model never contains a deferred source whose session has already closed. Use `EmailStoreSession.ReadItem` with
`preferStreamingAttachmentContent: true` when the caller wants on-demand payload streams.

Set `PstPassword` only when a protected PST requires checksum validation. Passwords are not logged or copied into results. Caller-owned streams must be readable and seekable; reads restore the original stream position and leave the stream open.

## Boundaries

- This package owns store/container traversal, bounded selection, validation/recovery discovery, new Unicode PST creation, and export/conversion orchestration.
- `OfficeIMO.Email` owns EML/MIME, MSG/OFT, TNEF, mbox, MAPI models, and item serialization.
- `OfficeIMO.Reader.EmailStore` owns optional Reader registration and rich-result projection.
- Existing PST/OST mutation, append, compaction, repair, password/encryption authoring, Outlook profiles, Exchange synchronization, cloud download, and server-side recovery remain outside this package.
- The writer currently emits Unicode PST files. It does not emit ANSI PST or OST files, and one data tree is limited to 4 GiB.

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.
- External runtime parser/writer dependencies: none.
- Direct first-party dependencies: `OfficeIMO.Email` and `OfficeIMO.Rtf`.
- `OfficeIMO.Email` carries Microsoft's `System.Text.Encoding.CodePages` compatibility package for legacy message
  encodings; no Outlook installation, native component, or third-party PST/OST/OLM parser is used.

See the [complete OfficeIMO package map](../README.md) for related formats and Reader adapters.
