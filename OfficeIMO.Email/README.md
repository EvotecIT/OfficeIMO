# OfficeIMO.Email

`OfficeIMO.Email` reads and writes persisted email and Outlook artifacts without MsgKit, MsgReader, OpenMcdf, RtfPipe, MimeKit, MailKit, or platform UI packages.

The package supports:

- EML and MIME messages, including multipart bodies, encoded headers, inline resources, attachments, and embedded messages
- Outlook MSG files with standard and named MAPI properties, legacy code pages, recipients, embedded messages, linked attachments, and OLE/custom-storage attachments
- MS-OXRTFCP compressed and uncompressed RTF bodies, including bounded expansion and checksum validation
- Outlook messages, appointments, contacts, tasks, journals, and sticky notes with typed read/write models
- TNEF payloads such as `winmail.dat`
- mboxo and mboxrd mailbox archives
- bounded synchronous and asynchronous reads, immutable reader configuration, cancellation, and structured diagnostics
- deterministic EML, MSG, TNEF, and mbox writing

MSG output includes the root storage identity, complete named-property forward and reverse mappings, and the compatibility metadata required by native Outlook. A document without an explicit date receives a deterministic creation-time fallback; set `Date` or `MessageMetadata.CreatedDate` when the real timestamp matters.

The package depends on `OfficeIMO.Rtf`, which owns RTF syntax and semantic conversion, and that package uses Microsoft's code-page compatibility library. MSG compound storage is shared OfficeIMO source compiled into `OfficeIMO.Email`; it is not a public general-purpose CFB package. This keeps the product graph small without copying RTF or CFB logic into the facade.

## Read and write a message

```csharp
using OfficeIMO.Email;

var options = new EmailReaderOptions(
    maxInputBytes: 64L * 1024L * 1024L,
    maxAttachmentBytes: 32L * 1024L * 1024L,
    includeAttachmentContent: true);

var reader = new EmailDocumentReader(options);
EmailReadResult result = await reader.ReadAsync("message.msg");

Console.WriteLine(result.Document.Subject);
Console.WriteLine(result.Document.Body.Text);

foreach (EmailDiagnostic diagnostic in result.Diagnostics) {
    Console.WriteLine($"{diagnostic.Severity}: {diagnostic.Code}: {diagnostic.Message}");
}

var writer = new EmailDocumentWriter();
writer.Write(result.Document, "message.eml", EmailFileFormat.Eml);
```

The same model can be written as EML, MSG, or TNEF. Conversion does not require a separate converter API:

```csharp
EmailDocument document = new EmailDocumentReader().Read("source.eml").Document;
new EmailDocumentWriter().Write(document, "converted.msg", EmailFileFormat.OutlookMsg);
```

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

## Outlook item projections

MSG and TNEF properties remain available through `MapiProperties` even when no convenience property exists. Common Outlook item fields are also projected onto typed models:

```csharp
EmailDocument item = new EmailDocumentReader().Read("meeting.msg").Document;

if (item.OutlookItemKind == OutlookItemKind.Appointment && item.Appointment is not null) {
    Console.WriteLine(item.Appointment.Start);
    Console.WriteLine(item.Appointment.End);
    Console.WriteLine(item.Appointment.Location);
}
```

Equivalent projections are available through `Contact`, `Task`, `Journal`, and `Note`.

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

## Protected Outlook messages

`EmailDocument.Protection` detects opaque and clear-signed S/MIME Outlook message classes and points to the original `.p7m` or `.p7s` attachment. It does not validate or decrypt that payload. Applications can pass `Protection.PayloadAttachment.Content` to MimeKit or another cryptographic owner when attachment content was retained.

```csharp
EmailDocument protectedMessage = new EmailDocumentReader().Read("signed.msg").Document;

if (protectedMessage.Protection.IsProtected) {
    byte[]? cms = protectedMessage.Protection.PayloadAttachment?.Content;
    // Mailozaurr or another host can process cms with its configured MimeKit context.
}
```

## Resource limits

`EmailReaderOptions` is immutable and applies limits before retaining decoded content. It controls source size, header size and count, MIME part count and depth, per-attachment and aggregate attachment bytes, embedded-message depth, CFB directory entries, MAPI properties and decoded property bytes, and TNEF attributes.

Use `includeAttachmentContent: false` when only attachment metadata is needed. Parsing still validates the source structure, but decoded attachment payloads are not retained in the result model.

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

The package does not expose general-purpose CFB transactions and does not read PST or OST mailbox stores. Its compound implementation serves MSG and structured attachments only. Use `EmailMailboxReader` for mbox aggregates and a dedicated store API for PST/OST workflows.

For an exact pass-through of a signed or encrypted artifact, read with `preserveRawSource: true` and write with `usePreservedRawSource: true`. A structured rewrite can change MIME canonicalization and must not be treated as signature-preserving.
