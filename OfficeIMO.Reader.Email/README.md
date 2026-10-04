# OfficeIMO.Reader.Email

One Reader package for the complete `OfficeIMO.Email` data surface:

Email body, text attachment, iCalendar, and vCard chunks retain complete Unicode scalars at `MaxChars` boundaries.
Text attachments use the shared email charset decoder before projection or delegation to a registered text handler.
Decoding recovery warnings appear in the document diagnostics and attachment chunk warnings.

Each ordinary attachment reports a Reader extraction outcome: succeeded, unsupported, content unavailable,
empty, or failed. Unsupported payloads are skipped before opening their content stream. Outcomes carry
the logical attachment path in document diagnostics and `ReaderEmailStoreItemResult.ItemDiagnostics`;
skip, empty and failure details also appear on attachment chunks. Document metadata includes attempted,
succeeded, skipped, empty and failed counts. Embedded messages are projected recursively and counted
separately. Successful extraction means readable content was produced; nested handler warnings still apply.

HTML and RTF bodies use the existing HTML adapter for semantic Markdown even when the host registers only
email handlers. A host's registered HTML handler takes precedence. Projection failures retain the safe HTML
source and report `EMAIL_BODY_READER_FAILED`.

Store attachments use bounded session streams by default, including OLM and EMLX.
Reader consumes supported attachment text before closing the session and returns asset metadata without
`PayloadBytes` for streamed content. To retain available attachment bytes in a result, register the store handler
with `new ReaderEmailStoreOptions { StreamAttachmentContent = false }` and keep the item's explicit streaming
preference disabled. `StoreOptions.RetainAttachmentContent = false` omits payloads in either mode.
Registered store safety limits, including `MaxDirectoryEntryCount`, survive option cloning; `MaxItems` bounds
the projected selection independently of the store's catalog limits.

Direct message, mailbox, iCalendar, and vCard handlers expose native asynchronous path and stream entry points.
`ReadDocumentAsync` uses their owning libraries' async I/O and Reader.Core's asynchronous source hashing.
Source hashes are computed once by Reader.Core when `ComputeHashes` is enabled, and caller streams remain open
with their original position restored.

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
