---
title: "Read, write, and convert email messages"
description: "Use .NET to read and convert EML, MIME, MSG, OFT, TNEF, winmail.dat, attachments, embedded messages, and Outlook item properties."
meta.seo_title: "Convert EML, MSG, OFT, TNEF, and winmail.dat in .NET"
order: 37
---

`EmailDocument` provides one model for EML/MIME, Outlook MSG and OFT, and TNEF payloads such as `winmail.dat`. It preserves recipients, bodies, attachments, inline resources, embedded messages, standard and named MAPI properties, and typed Outlook item fields.

## Load and save by extension

```csharp
using OfficeIMO.Email;

EmailDocument message = EmailDocument.Load("source.msg");

Console.WriteLine(message.Subject);
Console.WriteLine(message.Body.Text);

message.Save("converted.eml");
```

`Save` selects EML, MIME, MSG, OFT, or TNEF from the destination name. Use an explicit format when the destination has a neutral extension:

```csharp
message.Save("artifact.bin", EmailFileFormat.OutlookMsg);
```

The asynchronous API performs asynchronous source and destination I/O:

```csharp
EmailDocument message = await EmailDocument.LoadAsync(
    "source.eml",
    cancellationToken: cancellationToken);

await message.SaveAsync("converted.msg", cancellationToken: cancellationToken);
```

## Analyze conversion before writing

Known semantic loss is blocked by default. Analyze a conversion when your application needs to show the user why a destination cannot represent part of the source:

```csharp
var writer = new EmailDocumentWriter();
EmailConversionReport report =
    writer.AnalyzeConversion(message, EmailFileFormat.Eml);

foreach (EmailDiagnostic diagnostic in report.Diagnostics) {
    Console.WriteLine(
        $"{diagnostic.Severity}: {diagnostic.Code}: {diagnostic.Message}");
}

if (report.CanWrite) {
    writer.Write(message, "converted.eml");
}
```

Examples of reported boundaries include signed or encrypted content, Outlook recurrence and time-zone blobs, opaque Exchange contact identities, and item types with no standard EML representation. `EmailConversionLossPolicy.Warn` is available only for applications that have deliberately accepted the reported loss.

## Stream large attachments

Use the streaming reader when attachment content should not be retained in managed byte arrays:

```csharp
var options = new EmailReaderOptions(
    maxInputBytes: 64L * 1024L * 1024L,
    maxAttachmentBytes: 32L * 1024L * 1024L,
    includeAttachmentContent: true);

using EmailReadResult result = await new EmailDocumentReader(options)
    .ReadStreamingAsync("large-message.msg", cancellationToken);

foreach (EmailAttachment attachment in result.Document.Attachments) {
    using Stream content =
        await attachment.OpenContentStreamAsync(cancellationToken);
    await content.CopyToAsync(destination, 81920, cancellationToken);
}
```

Dispose the result after consuming file-backed content. Its temporary sources belong to the result and are removed when it is disposed.

## Outlook item projections

MSG and TNEF can represent appointments, contacts, tasks, journals, sticky notes, meeting lifecycle messages, and personal distribution lists. Common fields are projected onto typed models while unmodeled properties remain available in the MAPI property bag.

```csharp
EmailDocument item = EmailDocument.Load("meeting.msg");

if (item.OutlookItemKind == OutlookItemKind.Appointment &&
    item.Appointment is not null) {
    Console.WriteLine(item.Appointment.Start);
    Console.WriteLine(item.Appointment.End);
    Console.WriteLine(item.Appointment.Location);
}
```

Continue with [calendars and contacts](/docs/email/calendars-and-contacts/) for recurrence, ICS, and vCard workflows, or [safety and verification](/docs/email/safety-and-verification/) for diagnostics and migration checks.
