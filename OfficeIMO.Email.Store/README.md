# OfficeIMO.Email.Store

`OfficeIMO.Email.Store` is the fully managed, read-only mailbox-store package for PST, OST, OLM, and EMLX artifacts. It projects every supported item into the common `OfficeIMO.Email.EmailDocument` model without Outlook, native libraries, or third-party parser packages.

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

PST and OST MAPI properties use the same projections as MSG and OFT, so messages, appointments, contacts, tasks, journals, notes, recipients, attachments, and named properties do not acquire a second public model.

## Limits and attachment retention

Reads are bounded before large structures are retained. Applications can narrow source bytes, nodes, folders, items, properties, attachments, archive entries, XML characters, and embedded-message depth:

```csharp
var options = new EmailStoreReaderOptions(
    maxInputBytes: 2L * 1024 * 1024 * 1024,
    maxItemCount: 100_000,
    retainAttachmentContent: false,
    includeAssociatedItems: false);

EmailStoreReadResult result = new EmailStoreReader(options).Read("archive.ost");
```

Set `PstPassword` only when a protected PST requires checksum validation. Passwords are not logged or copied into results. Caller-owned streams must be readable and seekable; reads restore the original stream position and leave the stream open.

## Boundaries

- This package owns store/container traversal and store-to-`EmailDocument` projection.
- `OfficeIMO.Email` owns EML/MIME, MSG/OFT, TNEF, mbox, MAPI models, and item serialization.
- `OfficeIMO.Reader.EmailStore` owns optional Reader registration and rich-result projection.
- Store mutation, Outlook profiles, Exchange synchronization, cloud download, and server-side recovery are outside this offline reader.

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.
- External parser dependencies: none.
- OfficeIMO dependency: `OfficeIMO.Email`.

See the [complete OfficeIMO package map](../README.md) for related formats and Reader adapters.
