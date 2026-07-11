# OfficeIMO.Email support matrix

This matrix describes the current public contract for persisted email and Outlook artifacts. It separates format support from transport and cryptography so applications can combine `OfficeIMO.Email` with MailKit or MimeKit without pulling their concerns into the artifact engine.

## Formats and workflows

| Capability | Status | Current contract | Boundary |
| --- | --- | --- | --- |
| EML and MIME read/write | Broad | Headers, encoded words, common charsets, multipart bodies, plain text, HTML, inline resources, attachments, and embedded RFC messages | Mail transport, DKIM, ARC, PGP, and S/MIME cryptography belong to the host |
| Outlook MSG read/write | Broad | Standard and named MAPI properties, Unicode and String8 code pages, sender/representing/received addresses, all recipient roles, message metadata, bodies, attachments, embedded items, and deterministic output | PST/OST stores and Exchange directory resolution are not MSG artifact concerns |
| MSG compound storage | Complete for MSG | FAT, MiniFAT, DIFAT, hierarchical storages, regular/mini streams, embedded messages, and OLE/custom attachment storages | Not a public general-purpose CFB transaction API |
| Outlook messages | Broad | Subject normalization, conversation metadata, importance, priority, read/draft/receipt state, categories, reactions payloads, received-by metadata, and retained custom properties | Unknown or vendor properties remain available through `MapiProperties` |
| Appointments and meetings | Broad | Start/end, location, all-day, busy/meeting/response state, attendees, recurrence payloads, reminders, client intent, and time-zone payloads | Recurrence and time-zone binary structures are retained, not rewritten into a second calendar engine |
| Contacts | Broad | Names, company/department, dates, addresses, phones, three email slots, web/IM fields, privacy, picture metadata, and retained custom properties | Outlook distribution lists and directory lookups remain raw MAPI unless a typed contract is added |
| Tasks | Broad | Dates, status, completion, effort, owner/assignment state, reminders, ordering, contacts, companies, billing, and mileage | Server-side task synchronization is outside the artifact engine |
| Journals and sticky notes | Broad | Journal timing/type/document flags and note color/size/position | UI rendering is outside scope |
| RTF bodies | Broad | MS-OXRTFCP `LZFu` and `MELA`, bounded decompression, CRC diagnostics, encapsulated HTML projection, and deterministic writing | RTF syntax and semantic conversion are owned by `OfficeIMO.Rtf` |
| TNEF / `winmail.dat` | Broad | Message attributes, MAPI properties, recipient rows, attachments, embedded items, checksums, limits, and deterministic writing | Transport generation policy remains with the mail client |
| mboxo and mboxrd | Supported | Aggregate read/write, envelope metadata, escaping, message-count and source limits | No mailbox indexing or concurrent store engine |
| Protected Outlook messages | Handoff | Detects opaque and clear-signed S/MIME classes and exposes the original `.p7m`/`.p7s` payload attachment | Verification, trust, certificate/key lookup, and decryption belong to MimeKit or another host provider |
| Lossless pass-through | Supported | Preserved raw source can be emitted unchanged when explicitly requested | Structured edits regenerate the artifact and cannot preserve an existing cryptographic signature |
| OfficeIMO.Reader integration | Supported | Content detection, rich result projection, assets, embedded messages, chunks, and structured diagnostics | Reader remains a thin consumer of `OfficeIMO.Email` |

## MsgKit, MsgReader, and OpenMcdf replacement map

| Previous dependency capability | OfficeIMO owner |
| --- | --- |
| MsgKit EML-to-MSG and MSG authoring | `EmailDocumentReader`, `EmailDocumentWriter`, and the typed Outlook models |
| MsgKit sender, representing sender, recipient, body, metadata, and attachment builders | `EmailDocument`, `EmailAddress`, `EmailRecipient`, `EmailBody`, `EmailMessageMetadata`, and `EmailAttachment` |
| MsgReader MSG projection | `EmailDocumentReader` and typed message/appointment/contact/task/journal/note models |
| MsgReader unknown property access | Retained `MapiProperties` plus `GetMapiProperty` and `GetMapiValue` helpers |
| MsgReader compressed RTF and RTF-to-HTML | `OfficeIMO.Email` transport compression plus `OfficeIMO.Rtf` semantic projection |
| MsgReader signature parsing | Protected-payload handoff to the host's MimeKit cryptography policy |
| MsgReader nested MSG and `winmail.dat` expansion | `EmailAttachment.EmbeddedDocument` for embedded MSG, RFC message, and encapsulated TNEF content |
| OpenMcdf storage used by MSG | Internal OfficeIMO shared compound reader/writer compiled into `OfficeIMO.Email` |
| General-purpose OpenMcdf transactions | Intentionally not replaced by `OfficeIMO.Email`; consumers needing arbitrary CFB use a dedicated CFB library |

## Interoperability evidence

| Oracle or evidence | Result |
| --- | --- |
| MsgKit 3.0.5 runtime generation | OfficeIMO reads MsgKit EML-to-MSG output and named contact properties without mapping diagnostics |
| MsgReader 6.0.12 | OfficeIMO output is readable for message, recipient, room/resource, body, attachment, appointment, contact, task, and journal contracts |
| MsgReader 6.0.12 sample corpus | 15 real MSG fixtures matched subject, attachment count, and recipient count with no MSG parse errors, named-property warnings, or property-stream alignment warnings |
| OpenMcdf 3.1.4 | Test-only oracle opens OfficeIMO mini-stream, regular-stream, hierarchical, empty-stream, and DIFAT compound output |
| MimeKit 4.x TNEF reader | Accepts OfficeIMO TNEF output as compliant |
| Microsoft Outlook for Mac | Opens OfficeIMO-authored message, appointment, contact, task, journal, and note MSG files by their native subjects; the message view showed sender, recipient, body, and attachment content |
| Local packed-package consumer | A clean net8 consumer restored local `OfficeIMO.Email 0.1.0` and `OfficeIMO.Rtf 0.1.10`, wrote an MSG, and read it asynchronously |
| Performance contracts | Release tests cover 1 MiB MIME, 1 MiB MSG attachment, and 500-message mbox workloads; see [performance evidence](officeimo.email-performance.md) |

## Explicit non-goals

- SMTP, IMAP, POP3, Graph, authentication, and account synchronization
- DKIM, ARC, PGP, certificate trust, S/MIME verification, and decryption
- PST or OST mailbox stores
- a public arbitrary-CFB editing or transaction package
- Outlook UI automation or identical editors across platforms; Outlook for Mac uses its generic item viewer for non-mail MSG classes
- pretending that every vendor-specific named property has a typed convenience field; retained MAPI values are the compatibility escape hatch

Mailozaurr's migration and ownership split are documented in [Moving Mailozaurr MSG support to OfficeIMO.Email](officeimo.email-mailozaurr-migration.md).
