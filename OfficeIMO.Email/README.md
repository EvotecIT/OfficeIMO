# OfficeIMO.Email

`OfficeIMO.Email` reads and writes persisted email and Outlook artifacts without adding third-party runtime dependencies.

The package supports:

- EML and MIME messages, including multipart bodies, encoded headers, inline resources, attachments, and embedded messages
- Outlook MSG files with typed MAPI properties, named properties, recipients, embedded messages, and OLE/custom-storage attachments
- Outlook messages, appointments, contacts, tasks, journals, and sticky notes
- TNEF payloads such as `winmail.dat`
- mboxo and mboxrd mailbox archives
- bounded synchronous and asynchronous reads, immutable reader configuration, cancellation, and structured diagnostics
- deterministic EML, MSG, TNEF, and mbox writing

The product assembly has no NuGet package dependencies. Its compound-file implementation is shared OfficeIMO source compiled into the package; it is not a separate public CFB package.

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

`OfficeIMO.Email` owns offline artifact parsing, serialization, and format-neutral Outlook data. It does not connect to mail servers, authenticate users, resolve certificates or keys, verify DKIM/ARC/PGP/S/MIME signatures, or decrypt protected messages. MailKit, MimeKit, and applications such as Mailozaurr remain the right owners for those operations.

For an exact pass-through of a signed or encrypted artifact, read with `preserveRawSource: true` and write with `usePreservedRawSource: true`. A structured rewrite can change MIME canonicalization and must not be treated as signature-preserving.

Compressed RTF bodies are deliberately routed through `OfficeIMO.Rtf`; until that integration lands, existing compressed-RTF MAPI properties are retained but a newly assigned `EmailBody.Rtf` value is not written into MSG.
