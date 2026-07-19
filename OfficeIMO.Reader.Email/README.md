# OfficeIMO.Reader.Email

One Reader package for the complete `OfficeIMO.Email` data surface:

- EML, MSG/OFT, TNEF, Mbox/MBX, iCalendar, and vCard artifacts
- PST, OST, OLM, EMLX, Maildir, and mailbox-directory sessions
- Outlook Offline Address Book files

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddEmailHandlers()
    .Build();

OfficeDocumentReadResult message = reader.ReadDocument("message.msg");
OfficeDocumentReadResult store = reader.ReadDocument("archive.pst");
```

Install only this package when Reader needs email data. It depends on `OfficeIMO.Reader.Core` and `OfficeIMO.Email`; store and address-book support do not add separate NuGet layers.
