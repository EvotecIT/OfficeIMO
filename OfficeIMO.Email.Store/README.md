# OfficeIMO.Email.Store

`OfficeIMO.Email.Store` is the fully managed mailbox-store package for PST, OST, OLM, EMLX, Mbox, Apple Mail,
and Maildir sources. Read sessions never modify their source. Selected items project into the common
`OfficeIMO.Email.EmailDocument` model, and applications can create, verify, resume, or merge a separate Unicode PST
or explicitly mutate an existing Unicode PST through a verified atomic-rewrite transaction without Outlook, native
libraries, or third-party parser packages.

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
| EMLX | One Apple Mail EMLX item, including its RFC 5322/MIME message and supported XML property-list metadata. Partial files report that external Apple Mail content may be absent. Standalone and directory writers produce length-prefixed EML plus optional Apple property-list metadata. |
| Mbox | Mboxo and mboxrd archives exposed through the common store session, folder, item, query, validation, and export contracts. |
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

EmailStoreExportReport maildir = session.ExportToNativeDirectory(
    "exported-mail/maildir",
    new EmailStoreNativeDirectoryExportOptions(
        EmailStoreNativeDirectoryFormat.Maildir,
        maxItems: 50_000));

EmailStoreExportReport appleMail = session.ExportToNativeDirectory(
    "exported-mail/apple-mail",
    new EmailStoreNativeDirectoryExportOptions(
        EmailStoreNativeDirectoryFormat.Emlx,
        maxItems: 50_000));
```

Output conversion uses `OfficeIMO.Email` and its explicit semantic-loss policy. Existing destinations are not
replaced by default. Per-item failures and fidelity warnings remain visible in the export report. Maildir export
creates `tmp`, `new`, and `cur`; when the destination file system cannot represent the `:2,` flag suffix, flags remain
in the preservation manifest and the report contains a warning. `EmailStoreEmlxWriter` is also public for writing
one EMLX artifact directly.

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
journals, notes, and messages do not acquire a PST-only property model. `AddFolder` can assign supported standard
roles such as Inbox, Sent Items, Outbox, Calendar, Contacts, Tasks, Drafts, and view folders; the writer emits their
store/default-folder EntryID properties so readers and Outlook do not have to guess from localized display names.

Set `failOnDataLoss: true` when a warning should prevent the final commit. Examples include an attachment whose
payload was not retained, structured-storage metadata without its original compound payload, or a named property
whose source Name-to-ID mapping was unavailable. `Diagnostics` identifies the affected item or property without
including message or attachment content.

Writer cardinality state is disk-backed: block/allocation maps, NBT/BBT records, folder table rows, data-tree
indexes, and conversion mappings do not grow as retained managed lists. Configure a checkpoint path for durable
resume after cancellation or process interruption:

```csharp
const string checkpoint = "project-archive.pst.checkpoint";

using (EmailStorePstWriter writer = EmailStorePstWriter.Create(
    "project-archive.pst",
    new EmailStorePstWriterOptions(
        checkpointPath: checkpoint,
        checkpointIntervalItems: 1_000))) {
    // Add folders and items. Disposing an incomplete writer preserves its latest checkpoint.
}

using EmailStorePstWriter resumed = EmailStorePstWriter.Resume(checkpoint);
// Continue from the integrity-checked committed item boundary, then call Complete().
```

Resume truncates writer-owned working files to the last committed checkpoint before accepting more items. Normal
completion removes the checkpoint and temporary journals. `Abandon` or `DeleteCheckpoint` removes only the exact
writer-owned artifacts associated with that checkpoint.

## Mutate an existing Unicode PST

`EmailStorePstMutationTransaction` is a separate, explicit API for changing an existing Unicode PST. It locks the
source, stages folder and item operations, builds a replacement through the same managed writer, reopens and
semantically verifies the complete result, optionally commits a byte-for-byte backup, and only then atomically
replaces the source:

```csharp
using EmailStorePstMutationTransaction mutation =
    EmailStorePstMutationTransaction.Open(
        "archive.pst",
        new EmailStorePstMutationOptions(
            backupPath: "archive.before-mutation.pst"));

string projectFolder = mutation.CreateFolder("Project", mutation.RootFolderId);
string newItem = mutation.AddItem(projectFolder, new EmailDocument {
    Subject = "Added through a verified rewrite",
    MessageClass = "IPM.Note"
});

EmailStoreItemReference existing = mutation.EnumerateItems().First();
mutation.MoveItem(existing.Id, projectFolder);

EmailStorePstMutationReport mutationReport = mutation.Commit();
string rewrittenItemId = mutationReport.ItemIdMap[newItem];
```

The transaction can create, rename, move, and recursively delete non-mandatory folders; add, replace, move, and
delete items; and move items between normal and associated contents. Disposing without `Commit()` leaves the source
byte-for-byte unchanged. A no-op commit does not rewrite it. The transaction also detects source length or timestamp
changes before replacement. It holds both the source read lock and a path-scoped cross-process OfficeIMO mutation
lock through staging and replacement. The final replacement remains an atomic filesystem operation; software that
does not participate in the OfficeIMO lock must coordinate its own simultaneous replacement of the same path.

This is a verified semantic rewrite, not in-place NDB allocation-map editing. PST folder and item identifiers change,
so the report returns old/transaction-local to rewritten ID mappings. ANSI PST and OST inputs are rejected. Mandatory
folders cannot be renamed, moved, or deleted. Dynamic search-folder definitions cannot be regenerated; unsupported
search or writer fidelity is a blocking diagnostic while `failOnDataLoss` remains at its safe default. Applications
may opt into a known lossy static projection with `failOnDataLoss: false`, but verification of the resulting semantic
contents remains enabled unless explicitly disabled. Standard source-identified default-folder roles are rewritten
and verified as EntryID-backed roles, not merely as matching names. Items cannot be added or moved into the
writer-owned search folder because its contents are computed rather than authored as ordinary rows.

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

Conversion verification is enabled by default. The writer produces a same-directory staging PST, reopens it, and
compares every written item under the semantic migration profile before committing the destination. With
`failOnDataLoss: true`, a mismatch leaves an existing destination and manifest unchanged. The default uses an
ephemeral HMAC key. An optional TSV manifest can use a caller-supplied HMAC key for repeatable auditing; it contains
ordinals, statuses, keyed digests, an aggregate digest, and keyed HMAC tokens for differing canonical paths—never
subjects, addresses, filenames, content, arbitrary named-property names, raw difference paths, or store IDs.

## Merge multiple stores into one PST

`MergeToPst` accepts PST, OST, OLM, standalone EMLX, and mailbox-directory sources through one reusable engine:

```csharp
EmailStorePstMergeReport report = EmailStoreConverter.MergeToPst(
    new[] {
        new EmailStoreMergeSource("current.ost", "Current mailbox"),
        new EmailStoreMergeSource("archive.pst", "Archive"),
        new EmailStoreMergeSource("Outlook for Mac.olm"),
        new EmailStoreMergeSource("Apple Mail")
    },
    "combined.pst",
    new EmailStorePstMergeOptions(
        folderMode: EmailStoreMergeFolderMode.SeparateSourceRoots,
        deduplicate: true,
        maxRetries: 2));
```

Folder modes preserve isolated source roots, merge equal folder paths, or flatten items. Deduplication uses a
bounded on-disk hash index and a keyed semantic content fingerprint; associated items remain in a separate domain.
Transient source-open and item-read I/O failures are retried. Permanently unreadable sources/items are reported
according to the continuation policy. A destination writer or dedup-index failure aborts the atomic merge instead
of continuing from uncertain state. Reports contain aggregate counts, source outcomes, bounded diagnostics, and
privacy-safe progress.

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

Set `PstPassword` only when a protected PST requires checksum validation. Passwords are not logged or copied into results. Protected PSTs remain readable after validation, but mutation rejects them because the managed writer cannot preserve the protection. Caller-owned streams must be readable and seekable; reads restore the original stream position and leave the stream open.

## Boundaries

- This package owns store/container traversal, bounded selection, validation/recovery discovery, Mbox session projection, EMLX/Maildir output, new Unicode PST creation, verified conversion/merge, and verified atomic rewrite mutation of existing Unicode PST files.
- `OfficeIMO.Email` owns EML/MIME, MSG/OFT, TNEF, aggregate mbox, iCalendar, vCard, MAPI models, and item serialization.
- `OfficeIMO.Reader.EmailStore` owns optional Reader registration and rich-result projection.
- ANSI PST mutation, OST mutation/writing, in-place PST append/editing, compaction, repair, password/encryption authoring, OLM authoring, Outlook profiles, Exchange synchronization, cloud download, and server-side recovery remain outside this package.
- The writer currently emits Unicode PST files. It does not emit ANSI PST or OST files, and one data tree is limited to 4 GiB.
- Synthetic always-on gates cover deterministic round trips, malformed MIME/TNEF/compound input, memory budgets,
  checkpoints, merge/deduplication, and scale. libpff and classic Outlook mount/read/remove gates are opt-in so
  normal builds remain dependency-free; private PST/OST corpus tests are also opt-in and retain aggregate evidence only.

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.
- External runtime parser/writer dependencies: none.
- Direct first-party dependencies: `OfficeIMO.Email` and `OfficeIMO.Rtf`.
- `OfficeIMO.Email` carries Microsoft's `System.Text.Encoding.CodePages` compatibility package for legacy message
  encodings; no Outlook installation, native component, or third-party PST/OST/OLM parser is used.

See the [complete OfficeIMO package map](../README.md) for related formats and Reader adapters.
