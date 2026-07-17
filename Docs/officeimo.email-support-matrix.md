# OfficeIMO.Email support matrix

This matrix describes the current public contract for persisted email and Outlook artifacts. It separates format support from transport and cryptography so applications can combine `OfficeIMO.Email` with MailKit or MimeKit without pulling their concerns into the artifact engine.

## Formats and workflows

| Capability | Status | Current contract | Boundary |
| --- | --- | --- | --- |
| EML and MIME read/write | Broad | Headers, encoded words, common charsets, multipart bodies, plain text, HTML, inline resources, attachments, embedded RFC messages, file-backed attachment reads, and chunked sync/async writes | Mail transport, DKIM, ARC, PGP, and S/MIME cryptography belong to the host |
| Outlook MSG/OFT read/write | Broad | Standard and named MAPI properties, Unicode and String8 code pages, sender/representing/received addresses, all recipient roles, message metadata, bodies, attachments, embedded items, selective compound streams, and deterministic chunked output | PST/OST stores and Exchange directory resolution are not MSG artifact concerns |
| MSG compound storage | Complete for MSG | FAT, MiniFAT, DIFAT, hierarchical storages, regular/mini streams, embedded messages, and OLE/custom attachment storages | Not a public general-purpose CFB transaction API |
| Outlook messages | Broad | Subject normalization, conversation metadata, importance, priority, read/draft/receipt state, categories, reactions payloads, received-by metadata, and retained custom properties | Unknown or vendor properties remain available through `MapiProperties` |
| Appointments and meetings | Broad | Start/end, location, all-day, busy/meeting/response state, attendees, recurrence payloads, reminders, client intent, and time-zone payloads | Recurrence and time-zone binary structures are retained, not rewritten into a second calendar engine |
| Contacts | Broad | Names, company/department, dates, addresses, phones, three email slots, web/IM fields, privacy, picture metadata, and retained custom properties | Outlook distribution lists and directory lookups remain raw MAPI unless a typed contract is added |
| Tasks | Broad | Dates, status, completion, effort, owner/assignment state, reminders, ordering, contacts, companies, billing, and mileage | Server-side task synchronization is outside the artifact engine |
| Journals and sticky notes | Broad | Journal timing/type/document flags and note color/size/position | UI rendering is outside scope |
| RTF bodies | Broad | MS-OXRTFCP `LZFu` and `MELA`, bounded decompression, CRC diagnostics, encapsulated HTML projection, and deterministic writing | RTF syntax and semantic conversion are owned by `OfficeIMO.Rtf` |
| TNEF / `winmail.dat` | Broad | Message attributes, MAPI properties, recipient rows, attachments, embedded items, file-backed attachment reads, streaming checksums, limits, and deterministic writing | Transport generation policy remains with the mail client |
| mboxo and mboxrd | Supported | Aggregate read/write, envelope metadata, escaping, message-count and source limits | No mailbox indexing or concurrent store engine |
| Standalone iCalendar / ICS | Read/write/mutate | One or more ordered `VCALENDAR` roots; arbitrary nested components; repeated, IANA, and `X-` properties; multivalued parameters; UTF-8 75-octet folding; bounded sync/async I/O; DATE, floating, UTC, and TZID-local temporal helpers; RRULE helpers; and structural/conformance validation | TZID values are retained without host-OS normalization. Legacy `.vcs` data is parsed and preserved through the generic model, but RFC 5545 validation continues to report non-2.0 constructs |
| Standalone vCard / VCF | Read/write/mutate | Ordered multi-card streams for vCard 2.1, 3.0, and 4.0; groups; repeats; grouped properties; multivalued and legacy parameters; RFC 6868; quoted-printable continuation; binary/data-URI values; extensions; text helpers; and version-specific validation | Directory lookup/synchronization and vendor semantics without a standard mapping remain host concerns; unknown data stays accessible through the content-line model |
| PST and OST stores | Selective read; new Unicode PST write; verified Unicode PST mutation | ANSI/Unicode PST and supported PST-compatible OST NDB variants; bounded folder catalogs, item references, summaries, selective item parts, deferred attachment streams, associated items, and orphan discovery. New PST output covers folders, typed items, recipients, attachments, embedded messages, named properties, and multi-valued MAPI properties. Existing unprotected Unicode PST transactions stage folder/item operations, verify a complete rewrite, optionally back up, and atomically replace | Mutation is a semantic rewrite whose IDs change, not in-place NDB editing. Password-protected PSTs remain readable with validation but are rejected for mutation because protection cannot be preserved. ANSI PST mutation, OST mutation/output, append, repair, password/encryption authoring, Exchange synchronization, and recovery of content never cached in an OST are outside the contract |
| OLM, EMLX, Mbox, and mailbox directories | Read plus selected native write | Bounded Outlook for Mac ZIP/XML archives, individual Apple Mail EMLX items, partial-content metadata, Mbox store sessions, Apple Mail trees, Maildir, and EML/MIME directory sessions. EMLX and Maildir directory output is atomic per item and includes preservation diagnostics/manifests | OLM opens into a bounded materialized model; mailbox directories remain lazy. OLM authoring is not implemented; Maildir flag suffixes fall back to the manifest on file systems that cannot represent them |
| Store search and validation | Supported | Metadata queries, resumable semantic body/recipient/attachment-name search, snippets, progress, special-folder roles, content-availability reporting, and bounded PST/OST CRC/signature/layout validation | Search is an offline scan, not an Outlook or Exchange index query; structural validation does not repair the source |
| Store export, verified conversion, and merge | Supported | Selected items to EML, MSG, OFT, TNEF, Maildir, or EMLX; atomically committed streaming mbox; staged semantic verification before committing a new Unicode PST; and multi-source PST/OST/OLM/EMLX/Mbox/mailbox-directory merge with folder modes, disk-backed keyed deduplication, retries, and bounded diagnostics | Search folders become static folders; unavailable OST/server content and unsupported attachment payloads are reported rather than invented |
| Semantic fingerprints | Supported | Versioned migration, strict, and deduplication profiles; streamed attachment hashing; optional HMAC-SHA-256; and value-free difference reports | A fingerprint is a comparison/audit primitive, not a cryptographic authenticity signature |
| Outlook OAB address books | Read-only, selective | Bounded component discovery; dynamic-schema v4 Full Details entries and distribution lists; shared address/contact/MAPI projections; raw property retention; resumable search; seeded CRC, framing, and full-decode validation | Display templates and v2/v3 components are inspection-only; compressed Exchange downloads, patches, directory synchronization, and mutation are outside the expanded-cache reader |
| Protected Outlook messages | Handoff | Detects opaque and clear-signed S/MIME classes and exposes the original `.p7m`/`.p7s` payload attachment | Verification, trust, certificate/key lookup, and decryption belong to MimeKit or another host provider |
| Lossless pass-through | Supported | Preserved raw source can be emitted unchanged when explicitly requested | Structured edits regenerate the artifact and cannot preserve an existing cryptographic signature |
| OfficeIMO.Reader integration | Supported | Built-in handling includes individual email artifacts plus `.ics`, `.vcs`, `.vcf`, and `.vcard`; `OfficeIMO.Reader.EmailStore` adds selective PST/OST/OLM/EMLX projection, and `OfficeIMO.Reader.EmailAddressBook` adds selective typed OAB entry chunks | Reader remains a thin consumer of `OfficeIMO.Email`, `OfficeIMO.Email.Store`, and `OfficeIMO.Email.AddressBook` |

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
| iCalendar/vCard contract suite | Read-edit-write-reopen coverage includes multiple calendars/cards, nested alarms, recurrence and temporal forms, scoped TZID validation, vCard 2.1/3.0/4.0, grouped/repeated fields, media values, legacy parameter quoting, quoted-printable continuation, RFC 6868, Unicode octet folding, and configured size/depth/count limits |
| Microsoft Outlook for Mac | Opens OfficeIMO-authored message, appointment, contact, task, journal, and note MSG files by their native subjects; the message view showed sender, recipient, body, and attachment content |
| Local packed-package consumer | A clean net8 consumer restored locally packed `OfficeIMO.Drawing`, `OfficeIMO.Rtf`, `OfficeIMO.Email`, and `OfficeIMO.Email.Store` 2.0.1 artifacts from an isolated cache, then exercised asynchronous streaming output, semantic fingerprints, Unicode PST creation/reopen, and multi-store merge |
| Performance contracts | Release tests cover 1 MiB MIME, 1 MiB MSG attachment, 500-message mbox, a file-backed 16 MiB attachment, a 2,000-message PST, and 100,000-entry disk-backed mapping/dedup indexes; see [performance evidence](officeimo.email-performance.md) |
| Large-store contracts | A virtual 64 GiB PST contract covers selective reads, deferred attachment I/O, content search, and structural validation under fixed source-read ceilings; an aggregate-only 22.4 GB OST run exercises extraction, search, calendar items, Reader projection, and bounded structural checks |
| Unicode PST writer | Synthetic round trips cover multi-block heaps and tables, typed MAPI properties, recipients, large attachments, embedded and associated items, named properties, checkpoint/resume, verified OST-to-new-PST conversion, and multi-format merge. Independent libpff and classic Outlook gates open only generated synthetic stores. |
| Existing Unicode PST mutation | Synthetic transactions cover mixed folder/item create, rename, move, replace, associated-state change, recursive delete, optional byte-identical backup, semantic reopen verification, ID mappings, no-op/disposal source preservation, ANSI rejection, and cycle/mandatory-folder guards |
| Outlook OAB cache | Generated v4 fixtures cover every supported property encoding, corruption and limits; aggregate-only validation of 18 private cache components decoded and fully validated all 8,049 declared entries with no retained directory data |

## Explicit non-goals

- SMTP, IMAP, POP3, Graph, authentication, and account synchronization
- DKIM, ARC, PGP, certificate trust, S/MIME verification, and decryption
- In-place PST NDB mutation/append, ANSI PST mutation, OST mutation/output, compaction, repair, or password/encryption authoring
- OLM authoring, proprietary DBX mailboxes, Outlook/Mac profile databases, autocomplete caches, search indexes, synchronization state, and other profile/cache formats outside the dedicated OAB owner
- xCal/xCard, jCal/jCard, JSCalendar, and JSContact public adapters. The shared content-line model remains the canonical ICS/VCF owner so future XML/JSON adapters can reuse it without duplicating calendar/contact semantics; JSCalendar and JSContact require separate mapping contracts rather than syntax substitution
- a public arbitrary-CFB editing or transaction package
- Outlook UI automation or identical editors across platforms; Outlook for Mac uses its generic item viewer for non-mail MSG classes
- pretending that every vendor-specific named property has a typed convenience field; retained MAPI values are the compatibility escape hatch

Mailozaurr's migration and ownership split are documented in [Moving Mailozaurr MSG support to OfficeIMO.Email](officeimo.email-mailozaurr-migration.md).
