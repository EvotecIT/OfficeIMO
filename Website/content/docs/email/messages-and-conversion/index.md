---
title: "Read and Convert Email Messages"
description: "Read EML, MSG, OFT, and TNEF (winmail.dat) messages, convert between them, check for conversion loss, stream large attachments, and render a message to PDF."
meta.seo_title: "Convert EML, MSG, OFT, TNEF, and winmail.dat in .NET"
order: 37
packages:
  - OfficeIMO.Email
  - OfficeIMO.Email.Image
cmdlets:
  - Open-OfficeEmail
  - Save-OfficeEmail
  - New-OfficeEmailReaderOptions
  - New-OfficeEmailWriterOptions
  - Convert-OfficeDocument
commands:
  - officeimo convert
  - officeimo agent inspect-email
limits: "Conversions that would drop signed or encrypted content, recurrence data, or another known semantic are blocked unless the caller sets a conversion-loss policy."
search_terms:
  - msg to eml
  - eml to msg
  - winmail.dat
  - read outlook msg
  - email to pdf
verified: 2026-11-14
---

`EmailDocument` is one model for EML/MIME, Outlook MSG and OFT, and TNEF payloads such as `winmail.dat`. It keeps recipients, bodies, attachments, inline resources, embedded messages, MAPI properties, and typed Outlook item fields. `OfficeIMO.Email.Image` adds PDF and image rendering of a message.

## Read a message

### C#

```csharp
using OfficeIMO.Email;

EmailDocument message = EmailDocument.Load("message.msg");

Console.WriteLine(message.Subject);
Console.WriteLine(message.From?.Address);
Console.WriteLine(message.Date);
Console.WriteLine(message.Body.Text ?? message.Body.Html);

foreach (EmailRecipient recipient in message.Recipients) {
    Console.WriteLine($"{recipient.Kind}: {recipient.Address.Address}");
}

foreach (EmailAttachment attachment in message.Attachments) {
    Console.WriteLine($"{attachment.FileName} ({attachment.Length} bytes)");
}
```

### PowerShell

```powershell
$message = Open-OfficeEmail -Path .\message.msg
$message.Subject
$message.Attachments | Select-Object FileName, Length
```

### CLI

Print a bounded summary of a message:

```sh
officeimo agent inspect-email message.msg
```

## Convert MSG to EML and back

`Save` picks EML, MSG, OFT, or TNEF from the destination extension. Pass an `EmailFileFormat` when the destination has a neutral extension.

### C#

```csharp
using OfficeIMO.Email;

EmailDocument message = EmailDocument.Load("message.msg");
message.Save("message.eml");

EmailDocument fromEml = EmailDocument.Load("message.eml");
fromEml.Save("message.bin", EmailFileFormat.OutlookMsg);
```

The asynchronous methods use asynchronous file I/O and observe cancellation:

```csharp
using OfficeIMO.Email;

using var cancellation = new CancellationTokenSource(TimeSpan.FromMinutes(1));

EmailDocument message = await EmailDocument.LoadAsync("message.eml",
    cancellationToken: cancellation.Token);
await message.SaveAsync("message.msg", cancellationToken: cancellation.Token);
```

### PowerShell

```powershell
Open-OfficeEmail -Path .\message.msg | Save-OfficeEmail -Path .\message.eml
Open-OfficeEmail -Path .\winmail.dat | Save-OfficeEmail -Path .\message.msg -PassThru
```

### CLI

```sh
officeimo convert message.msg message.eml
officeimo convert winmail.dat message.msg
```

## Check what a conversion would lose

Known semantic loss is blocked by default. `AnalyzeConversion` reports why a destination format cannot represent part of the source, without writing anything.

### C#

```csharp
using OfficeIMO.Email;

EmailDocument message = EmailDocument.Load("meeting.msg");

var writer = new EmailDocumentWriter();
EmailConversionReport report = writer.AnalyzeConversion(message, EmailFileFormat.Eml);

foreach (EmailDiagnostic diagnostic in report.NativeDiagnostics) {
    Console.WriteLine($"{diagnostic.Severity}: {diagnostic.Code}: {diagnostic.Message}");
}

if (report.CanWrite) {
    writer.Write(message, "meeting.eml", EmailFileFormat.Eml);
}
```

To write anyway after showing the user what is lost, use a writer with an explicit loss policy:

```csharp
using OfficeIMO.Email;

EmailDocument message = EmailDocument.Load("meeting.msg");

var writer = new EmailDocumentWriter(new EmailWriterOptions(EmailConversionLossPolicy.Warn));
EmailWriteResult result = writer.Write(message, "meeting.eml", EmailFileFormat.Eml);

foreach (EmailDiagnostic diagnostic in result.Diagnostics) {
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
}
```

### PowerShell

```powershell
$options = New-OfficeEmailWriterOptions -AllowLoss
Open-OfficeEmail -Path .\meeting.msg | Save-OfficeEmail -Path .\meeting.eml -Options $options -PassThru
```

## Stream large attachments

Read into temporary files when attachment content should not stay in memory. Dispose the result after use: its temporary files are removed with it.

### C#

```csharp
using OfficeIMO.Email;

var options = new EmailReaderOptions(
    maxInputBytes: 64L * 1024 * 1024,
    maxAttachmentBytes: 32L * 1024 * 1024);

using EmailReadResult result = await new EmailDocumentReader(options)
    .ReadDocumentAsync("large-message.msg", EmailReadContentStorage.TemporaryFiles);

Directory.CreateDirectory("attachments");
int index = 0;
foreach (EmailAttachment attachment in result.Document.Attachments) {
    await using Stream content = await attachment.OpenContentStreamAsync();
    await using FileStream output = File.Create(Path.Combine("attachments", $"{index++:D3}.bin"));
    await content.CopyToAsync(output);
}
```

Attachment names come from the message. Sanitize them before using them as file names.

### PowerShell

```powershell
$options = New-OfficeEmailReaderOptions -MaxInputBytes 64MB -MaxAttachmentBytes 32MB
$result = Open-OfficeEmail -Path .\large-message.msg -Options $options -AsResult
$result.Diagnostics | Format-Table Severity, Code, Message
```

## Read Outlook appointments and other items

MSG and TNEF can hold appointments, contacts, tasks, journals, sticky notes, meeting messages, and distribution lists. Common fields are projected onto typed models, and the rest stay in the MAPI property bag.

### C#

```csharp
using OfficeIMO.Email;

EmailDocument item = EmailDocument.Load("meeting.msg");

if (item.OutlookItemKind == OutlookItemKind.Appointment && item.Appointment is not null) {
    Console.WriteLine(item.Appointment.Start);
    Console.WriteLine(item.Appointment.End);
    Console.WriteLine(item.Appointment.Location);
}
```

## Render a message to PDF

The PDF has a header block (subject, sender, recipients, date), the selected body, inline images, and a list of file attachments. Remote resources are blocked by default.

### C#

```csharp
using OfficeIMO.Email;

EmailDocument message = EmailDocument.Load("message.eml");

await using FileStream pdf = File.Create("message.pdf");
await message.SaveAsPdfAsync(pdf, new EmailPdfExportOptions {
    IncludeAttachmentList = true
});
```

### PowerShell

```powershell
Convert-OfficeDocument -Path .\message.eml -OutputPath .\message.pdf
```

### CLI

```sh
officeimo convert message.msg message.pdf
```
