# OfficeIMO.Email

`OfficeIMO.Email` reads and writes persisted email and Outlook artifacts without MsgKit, MsgReader, OpenMcdf, RtfPipe, MimeKit, MailKit, or platform UI packages.

The package supports:

- EML and MIME messages, including multipart bodies, encoded headers, inline resources, attachments, and embedded messages
- Outlook MSG and OFT files with standard and named MAPI properties, legacy code pages, recipients, embedded messages, linked attachments, and OLE/custom-storage attachments
- MS-OXRTFCP compressed and uncompressed RTF bodies, including bounded expansion and checksum validation
- Outlook messages, appointments, contacts, tasks, journals, and sticky notes with typed read/write models
- standards-based iCalendar (`VEVENT`/`VTODO`) and vCard projection when appointments, tasks, or contacts cross the EML boundary
- TNEF payloads such as `winmail.dat`
- mboxo and mboxrd mailbox archives
- reopenable attachment content, file-backed streaming reads, and per-message mbox streaming for large or store-backed workflows
- bounded synchronous and genuinely asynchronous source/destination I/O, immutable reader configuration, cancellation, and structured diagnostics
- deterministic EML, MSG, OFT, TNEF, and mbox writing with explicit conversion-loss policy
- versioned semantic fingerprints and value-free comparison reports for migration verification and deduplication

MSG output includes the root storage identity, complete named-property forward and reverse mappings, and the compatibility metadata required by native Outlook. A document without an explicit date receives a deterministic creation-time fallback; set `Date` or `MessageMetadata.CreatedDate` when the real timestamp matters.

The package depends on `OfficeIMO.Rtf`, which owns RTF syntax and semantic conversion. `OfficeIMO.Email` directly declares Microsoft's code-page compatibility library because its MIME, MSG, and TNEF readers use legacy character encodings themselves. MSG compound storage is shared OfficeIMO source compiled into `OfficeIMO.Email`; it is not a public general-purpose CFB package. This keeps the product graph explicit without copying RTF or CFB logic into the facade.

## Read and write a message

```csharp
using OfficeIMO.Email;

EmailDocument message = EmailDocument.Load("message.msg");

Console.WriteLine(message.Subject);
Console.WriteLine(message.Body.Text);

message.Save("message.eml");
```

`Save` infers EML, MSG, OFT, or TNEF from `.eml`, `.mime`, `.msg`, `.oft`, `.tnef`, or `winmail.dat`. The convenience methods throw `InvalidDataException` instead of silently returning a partial document or output when an error diagnostic is produced. Use the explicit overload only when the destination name has another extension:

```csharp
message.Save("artifact.bin", EmailFileFormat.OutlookMsg);
```

The async API has the same shape:

```csharp
EmailDocument message = await EmailDocument.LoadAsync("message.msg", cancellationToken: cancellationToken);
await message.SaveAsync("message.eml", cancellationToken: cancellationToken);
```

Use `EmailDocumentReader` when the host needs custom limits, raw-source preservation, or structured diagnostics:

```csharp
var options = new EmailReaderOptions(
    maxInputBytes: 64L * 1024L * 1024L,
    maxAttachmentBytes: 32L * 1024L * 1024L,
    includeAttachmentContent: true);

EmailReadResult result = await new EmailDocumentReader(options).ReadAsync("message.msg");

foreach (EmailDiagnostic diagnostic in result.Diagnostics) {
    Console.WriteLine($"{diagnostic.Severity}: {diagnostic.Code}: {diagnostic.Message}");
}

if (result.HasErrors) {
    // The advanced API leaves recovery policy to the host.
}
```

When attachment payloads should not be retained in managed arrays, use the streaming reader. Ordinary MIME,
MSG/OFT, and TNEF attachments become reopenable temporary-file sources while headers, bodies, recipients, and typed
Outlook fields keep the same model:

```csharp
using EmailReadResult result = await new EmailDocumentReader(options)
    .ReadStreamingAsync("large-message.msg", cancellationToken);

foreach (EmailAttachment attachment in result.Document.Attachments) {
    using Stream content = await attachment.OpenContentStreamAsync(cancellationToken);
    await content.CopyToAsync(destination, 81920, cancellationToken);
}
```

Dispose the `EmailReadResult` after consuming its file-backed content. Its temporary files and attachment sources
belong to that result and become unavailable after disposal. `Write` and `WriteAsync` consume `ContentSource`
attachments in chunks; the async path uses asynchronous source and destination I/O rather than wrapping the
synchronous writer in a task.

The same model can be written as EML, MSG, OFT, or TNEF. Conversion is a load followed by a save:

```csharp
EmailDocument.Load("source.eml").Save("converted.msg");
```

Known semantic loss is blocked by default. This includes changing or cross-converting a signed/encrypted artifact,
dropping Outlook recurrence or time-zone blobs into iCalendar, converting an Exchange contact whose address depends
on an opaque entry identifier, and writing journal or sticky-note semantics as EML. Analyze the conversion before
opening a destination when the host needs to present those choices:

```csharp
var writer = new EmailDocumentWriter();
EmailConversionReport report = writer.AnalyzeConversion(message, EmailFileFormat.Eml);

foreach (EmailDiagnostic diagnostic in report.Diagnostics) {
    Console.WriteLine($"{diagnostic.Severity}: {diagnostic.Code}: {diagnostic.Message}");
}

if (report.CanWrite) {
    writer.Write(message, "converted.eml");
}
```

Use `EmailConversionLossPolicy.Warn` only when the application has made an explicit decision to accept a documented
loss. Opaque MAPI/TNEF metadata that has no portable EML equivalent is reported as a warning while common message
content remains writable.

## Compare message semantics

Semantic comparison hashes a canonical, versioned projection rather than serialized bytes. The migration profile
normalizes store identity and serialization details, the strict profile includes representation details, and the
deduplication profile excludes synchronization and modification state:

```csharp
EmailSemanticComparisonReport comparison = EmailSemanticComparer.Compare(source, destination);

if (!comparison.IsMatch) {
    foreach (EmailSemanticDifference difference in comparison.Differences) {
        Console.WriteLine($"{difference.Kind}: {difference.Path}");
    }
}
```

Difference reports contain canonical paths and lengths, not message values. Before persisting fingerprints for
private mail, provide a random `digestKey` through `EmailSemanticComparisonOptions`; the resulting HMAC-SHA-256
fingerprints cannot be correlated without that caller-owned key.

## Read an mbox archive

An mbox file is an aggregate, so it has a dedicated reader rather than pretending to be one message:

```csharp
var mailboxReader = new EmailMailboxReader();
EmailMailboxReadResult result = mailboxReader.Read("archive.mbox");

foreach (EmailMailboxEntry entry in result.Mailbox.Messages) {
    Console.WriteLine($"{entry.EnvelopeSender}: {entry.Document.Subject}");
}
```

`EmailMailboxWriter` writes mboxo or mboxrd with deterministic envelope and escaping behavior.

For a large mailbox, enumerate or write one entry at a time:

```csharp
var reader = new EmailMailboxReader();

foreach (EmailMailboxEntryReadResult item in reader.ReadEntries("archive.mbox")) {
    Console.WriteLine(item.Entry.Document.Subject);
}

using Stream destination = File.Create("export.mbox");
new EmailMailboxWriter().WriteEntries(entries, destination);
```

`ReadEntries` retains at most one decoded message and restores a seekable caller-owned stream when enumeration ends.
`WriteEntries` writes completed entries directly, so a later enumeration or size failure can leave earlier entries in
the destination; use the aggregate `EmailMailbox.Save` API when atomic file replacement is required.

## Outlook item projections

MSG and TNEF properties remain available through `MapiProperties` even when no convenience property exists. Common Outlook item fields are also projected onto typed models:

```csharp
EmailDocument item = EmailDocument.Load("meeting.msg");

if (item.OutlookItemKind == OutlookItemKind.Appointment && item.Appointment is not null) {
    Console.WriteLine(item.Appointment.Start);
    Console.WriteLine(item.Appointment.End);
    Console.WriteLine(item.Appointment.Location);
}
```

Equivalent projections are available through `Contact`, `Task`, `Journal`, and `Note`.

An Outlook template uses the same model and compound-file engine as MSG while retaining its template identity:

```csharp
EmailDocument template = EmailDocument.Load("meeting.oft");
template.Subject = "Reusable meeting request";
template.Save("updated.oft");
```

When an appointment or task is written as EML, it becomes a `text/calendar` iCalendar part. Contacts become vCard
attachments. Reminders become `VALARM`; task fields without a direct iCalendar property use valid `X-OFFICEIMO-*`
extensions so they survive an OfficeIMO EML/MSG/TNEF cycle while remaining ignorable to other calendar readers.
Reading those parts restores the corresponding typed model. Source calendar/vCard bytes are retained
while the projected model is unchanged; editing a projected item is blocked by default when regeneration could omit
unmodeled source properties. Meetings whose only attendee data is display text are also blocked because a valid
iCalendar `ATTENDEE` requires an address. Journal and sticky-note models remain lossless across MSG/TNEF, but have no
standard EML representation and therefore require an explicit non-blocking loss policy for EML output.

Properties that do not have a convenience field remain available through `MapiProperties`. Standard, numeric named, and string named properties have public lookup helpers:

```csharp
string? displayName = item.MapiProperties.GetMapiValue<string>(0x3001);
MapiProperty? custom = item.MapiProperties.GetMapiProperty(
    new Guid("00062008-0000-0000-C000-000000000046"),
    "CustomName");
```

## RTF bodies

`EmailBody.Rtf` contains the byte-preserving RTF source decoded from `PidTagRtfCompressed`. The MSG and TNEF writers serialize an assigned RTF body with the MS-OXRTFCP compression format, and the reader accepts both the `LZFu` and `MELA` forms.

RTF syntax editing and semantic conversion belong to `OfficeIMO.Rtf`. Generate RTF through that package when the source contains characters that need RTF escapes. `OfficeIMO.Reader` will route an RTF-only email body through the registered `OfficeIMO.Reader.Rtf` handler; without that optional adapter, it preserves the RTF source and reports the fallback explicitly.

## Protected messages

`EmailDocument.Protection` detects opaque and clear-signed S/MIME plus signed or encrypted OpenPGP/MIME wrappers. It
does not validate signatures or decrypt payloads. Protected input retains its original artifact bytes automatically.
Writing an unchanged protected document in its source format emits those bytes verbatim; an edited or cross-format
write is blocked by default because regenerating the wrapper would invalidate its cryptographic meaning.

```csharp
EmailDocument protectedMessage = EmailDocument.Load("signed.msg");

if (protectedMessage.Protection.IsProtected) {
    byte[]? cms = protectedMessage.Protection.PayloadAttachment?.Content;
    // A cryptographic owner can validate or decrypt the retained payload.
}
```

## Mailbox stores and store-backed content

PST, OST, Outlook for Mac OLM, and Apple Mail EMLX containers belong to the optional `OfficeIMO.Email.Store` package. It yields ordinary `EmailDocument` instances while preserving folder paths, store metadata, diagnostics, and bounded attachment behavior:

```csharp
using OfficeIMO.Email.Store;

using EmailStoreSession session = EmailStoreSession.Open("archive.pst");
EmailStoreItemReference firstReference = session.EnumerateItems(
    new EmailStoreEnumerationOptions(maxItems: 1)).Single();
EmailDocument firstMessage = session.ReadItem(firstReference).Document;
```

`OfficeIMO.Email` remains sufficient for individual EML, MSG, OFT, TNEF, and mbox artifacts.

### Store-backed attachment content

`EmailAttachment.Content` remains the simple in-memory representation. A mailbox or archive provider can instead set
`ContentSource` to an `IEmailContentSource` that opens a fresh decoded stream on demand:

```csharp
EmailAttachment attachment = item.Attachments[0];

using Stream content = await attachment.OpenContentStreamAsync(cancellationToken);
await content.CopyToAsync(destination, 81920, cancellationToken);
```

EML, MSG, and TNEF writers consume either representation through the same bounded path. The interface deliberately
contains no PST, OST, MAPI, or Outlook types, so mailbox and store owners can yield ordinary `EmailDocument`
instances while keeping source lifetime and property-stream access in the owning package.

## Resource limits

`EmailReaderOptions` is immutable and applies limits before retaining decoded content. It controls source size, header size and count, MIME part count and depth, per-attachment and aggregate attachment bytes, embedded-message depth, CFB directory entries, MAPI properties and decoded property bytes, and TNEF attributes.

Use `includeAttachmentContent: false` when only attachment metadata is needed. Parsing still validates the source
structure, but ordinary decoded attachment payloads are not retained in the result model. Calendar and vCard parts
that define the typed item are retained because they are semantic message content, not optional file payloads.

## Reader integration

`OfficeIMO.Reader` recognizes `.eml`, `.msg`, `.oft`, `.mbox`, `.mbx`, `.tnef`, and `winmail.dat`. Add `OfficeIMO.Reader.EmailStore` for PST, OST, OLM, and EMLX. Its rich result includes envelope and Outlook metadata, structured diagnostics, materializable attachment assets, embedded messages, and chunks extracted from supported attachment formats.

```csharp
using OfficeIMO.Reader;

OfficeDocumentReadResult result = OfficeDocumentReader.Default.ReadDocument("message.msg");

Console.WriteLine(result.Source.Title);
foreach (OfficeDocumentAsset attachment in result.Assets) {
    Console.WriteLine($"{attachment.FileName}: {attachment.LengthBytes}");
}
```

## Scope boundary

`OfficeIMO.Email` owns offline artifact parsing, serialization, and format-neutral Outlook data. It does not connect to mail servers, authenticate users, resolve certificates or keys, verify DKIM/ARC/PGP/S/MIME signatures, or decrypt protected messages. MailKit, MimeKit, and applications such as Mailozaurr remain the owners for those operations.

The package does not expose general-purpose CFB transactions or mailbox-store traversal. Its compound implementation
serves MSG/OFT and structured attachments only. `OfficeIMO.Email.Store` is the separate source and new-PST owner
for PST, OST, OLM, EMLX, Apple Mail, and Maildir traversal, selection, validation, export, verified conversion, and
multi-store merge. Existing PST/OST mutation, append, repair, compaction, and OST writing remain outside both
packages.

For exact pass-through of an ordinary unprotected artifact, read with `preserveRawSource: true` and write with
`usePreservedRawSource: true`. Protected artifacts use safe unchanged pass-through automatically.

## Dependency footprint

- **External:** No third-party email engine or Outlook interop. `System.Text.Encoding.CodePages` supplies legacy encodings.
- **OfficeIMO:** `OfficeIMO.Drawing` and `OfficeIMO.Rtf`. MIME, MSG/MAPI, TNEF, mbox, and compressed-RTF handling are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
