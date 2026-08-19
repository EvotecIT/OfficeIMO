# OfficeIMO.Reader.Email

One Reader package for the complete `OfficeIMO.Email` data surface:

- EML, MSG/OFT, TNEF, Mbox/MBX, iCalendar, and vCard artifacts
- MHT/MHTML web archives with embedded MIME resources projected through `OfficeIMO.Reader.Html`
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
OfficeDocumentReadResult webArchive = reader.ReadDocument("snapshot.mhtml");
```

Register only MHTML when the other email handlers are not needed:

```csharp
OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddMhtmlHandler()
    .Build();
```

Install this package when Reader needs email or MHTML data. It depends on `OfficeIMO.Reader.Core`, `OfficeIMO.Email`, `OfficeIMO.Email.Html`, `OfficeIMO.Mhtml`, and the lean `OfficeIMO.Reader.Html` projection. Email HTML/text/Markdown preparation reuses the shared safe body and embedded-resource contract; store and address-book support do not add separate NuGet layers or another email model.
