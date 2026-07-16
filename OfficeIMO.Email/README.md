# OfficeIMO.Email

`OfficeIMO.Email` reads and writes persisted email and Outlook artifacts without MsgKit, MsgReader, OpenMcdf, RtfPipe, MimeKit, MailKit, or platform UI packages.

The package supports:

- EML and MIME messages, including multipart bodies, encoded headers, inline resources, attachments, and embedded messages
- Outlook MSG files with standard and named MAPI properties, legacy code pages, recipients, embedded messages, linked attachments, and OLE/custom-storage attachments
- MS-OXRTFCP compressed and uncompressed RTF bodies, including bounded expansion and checksum validation
- Outlook messages, appointments, contacts, tasks, journals, and sticky notes with typed read/write models
- standards-based iCalendar (`VEVENT`/`VTODO`) and vCard projection when appointments, tasks, or contacts cross the EML boundary
- TNEF payloads such as `winmail.dat`
- mboxo and mboxrd mailbox archives
- reopenable attachment content and per-message mbox streaming for large or store-backed workflows
- bounded synchronous and asynchronous reads, immutable reader configuration, cancellation, and structured diagnostics
- deterministic EML, MSG, TNEF, and mbox writing with explicit conversion-loss policy

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

`Save` infers EML, MSG, or TNEF from `.eml`, `.mime`, `.msg`, `.tnef`, or `winmail.dat`. The convenience methods throw `InvalidDataException` instead of silently returning a partial document or output when an error diagnostic is produced. Use the explicit overload only when the destination name has another extension:

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

The same model can be written as EML, MSG, or TNEF. Conversion is a load followed by a save:

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

## Store-backed attachment content

`EmailAttachment.Content` remains the simple in-memory representation. A mailbox or archive provider can instead set
`ContentSource` to an `IEmailContentSource` that opens a fresh decoded stream on demand:

```csharp
EmailAttachment attachment = item.Attachments[0];

using Stream content = await attachment.OpenContentStreamAsync(cancellationToken);
await content.CopyToAsync(destination, 81920, cancellationToken);
```

EML, MSG, and TNEF writers consume either representation through the same bounded path. The interface deliberately
contains no PST, OST, MAPI, or Outlook types, so a future mailbox-store package can yield ordinary `EmailDocument`
instances while keeping store lifetime and property-stream access in the store owner.

## Resource limits

`EmailReaderOptions` is immutable and applies limits before retaining decoded content. It controls source size, header size and count, MIME part count and depth, per-attachment and aggregate attachment bytes, embedded-message depth, CFB directory entries, MAPI properties and decoded property bytes, and TNEF attributes.

Use `includeAttachmentContent: false` when only attachment metadata is needed. Parsing still validates the source
structure, but ordinary decoded attachment payloads are not retained in the result model. Calendar and vCard parts
that define the typed item are retained because they are semantic message content, not optional file payloads.

## Reader integration

`OfficeIMO.Reader` recognizes `.eml`, `.msg`, `.mbox`, `.mbx`, `.tnef`, and `winmail.dat`. Its rich result includes envelope and Outlook metadata, structured diagnostics, materializable attachment assets, embedded messages, and chunks extracted from supported attachment formats.

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

The package does not expose general-purpose CFB transactions and does not yet read or modify PST/OST mailbox stores.
Its compound implementation serves MSG and structured attachments only. `IEmailContentSource`, the format-neutral
`EmailDocument`, and streaming mbox APIs are the compatibility boundary for a future dedicated PST/OST package; store
folder traversal, named-property mapping, item mutation, allocation tables, and transactional commits still belong in
that separate owner.

For exact pass-through of an ordinary unprotected artifact, read with `preserveRawSource: true` and write with
`usePreservedRawSource: true`. Protected artifacts use safe unchanged pass-through automatically.

## Dependency footprint

- **External:** No third-party email engine or Outlook interop. `System.Text.Encoding.CodePages` supplies legacy encodings.
- **OfficeIMO:** `OfficeIMO.Drawing` and `OfficeIMO.Rtf`. MIME, MSG/MAPI, TNEF, mbox, and compressed-RTF handling are first-party.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
