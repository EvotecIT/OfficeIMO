# OfficeIMO.Email

`OfficeIMO.Email` provides a first-party engine for persisted email and Outlook artifacts without a third-party message, compound-file, MIME, RTF, or platform-UI runtime.

```powershell
dotnet add package OfficeIMO.Email
```

S/MIME structure detection and unchanged protected-message pass-through are included without a cryptographic package.
Install the provider only when the application signs, encrypts, verifies, or decrypts S/MIME:

```powershell
dotnet add package OfficeIMO.Security
```

Email reads expose an aggregate `ProcessingBudget` snapshot. MIME, MSG, embedded TNEF, and streaming payload paths share one operation ledger, so nested parsers do not receive fresh attachment, property, or structural allowances. `EmailLimitExceededException.Diagnostic` carries a stable actionable non-sensitive diagnostic automatically; store limits expose the equivalent through `EmailStoreDiagnostic.FromLimit`.

Store table, content-search, and OAB search checkpoints are versioned persistence values bound to a complete-source SHA-256 and exact query signature. Persist the checkpoint's `Value`, parse it after restart, and expect resume to fail closed if the source bytes or query changed. Creating a durable checkpoint intentionally reads the complete selected source to establish that identity.

`EmailStoreSession.PlanMaintenance` is read-only and source-bound. It returns `CompleteInspection` rather than `None` whenever item, recovery, structural page, block, or byte bounds leave evidence incomplete, or when requested structural verification is unsupported for the source format. Executable repair remains a separate operation through recovery export, PST compaction, or PST split planning; OfficeIMO never repairs the opened source in place and rewrite paths require semantic post-verification.

The package supports:

- EML and MIME messages, including multipart bodies, encoded headers, inline resources, attachments, and embedded messages
- Outlook MSG and OFT files with standard and named MAPI properties, legacy code pages, recipients, embedded messages, linked attachments, and OLE/custom-storage attachments
- MS-OXRTFCP compressed and uncompressed RTF bodies, including bounded expansion and checksum validation
- Outlook messages, appointments, contacts, tasks, journals, and sticky notes with typed read/write models
- standalone iCalendar (`.ics`) and vCard (`.vcf`/`.vcard`) documents with ordered, lossless content-line mutation, validation, bounded I/O, and standards-based EML projection
- TNEF payloads such as `winmail.dat`
- mboxo and mboxrd mailbox archives
- PST, OST, Outlook for Mac OLM, EMLX, Apple Mail, Maildir, and mailbox-directory stores
- Outlook Offline Address Book discovery, v4 entry decoding, search, validation, and offline identity resolution
- one mixed-artifact discovery API for individual messages, calendars, contacts, stores, and OAB files
- reopenable attachment content, file-backed streaming reads, and per-message mbox streaming for large or store-backed workflows
- bounded synchronous and genuinely asynchronous source/destination I/O, immutable reader configuration, cancellation, and structured diagnostics
- deterministic EML, MSG, OFT, TNEF, and mbox writing with explicit conversion-loss policy
- versioned semantic fingerprints and value-free comparison reports for migration verification and deduplication

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

`Save` infers EML, MSG, OFT, or TNEF from `.eml`, `.mime`, `.msg`, `.oft`, `.tnef`, or `winmail.dat`. The convenience methods throw `InvalidDataException` instead of silently returning a partial document or output when an error diagnostic is produced. Use the explicit overload only when the destination name has another extension:

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

When attachment payloads should not be retained in managed arrays, use the streaming reader. Ordinary MIME,
MSG/OFT, and TNEF attachments become reopenable temporary-file sources while headers, bodies, recipients, and typed
Outlook fields keep the same model:

```csharp
using EmailReadResult result = await new EmailDocumentReader(options)
    .ReadStreamingAsync("large-message.msg", cancellationToken);

foreach (EmailAttachment attachment in result.Document.Attachments) {
    using Stream content = await attachment.OpenContentStreamAsync(cancellationToken);
    await content.CopyToAsync(destination, 81920, cancellationToken);
}
```

Dispose the `EmailReadResult` after consuming its file-backed content. Its temporary files and attachment sources
belong to that result and become unavailable after disposal. `Save`, `SaveAsync`, `Write`, and `WriteAsync` consume
`ContentSource` attachments through the same bounded writer pipelines instead of first materializing the complete
artifact. The async path uses asynchronous source and destination I/O rather than wrapping the synchronous writer
in a task.

The same model can be written as EML, MSG, OFT, or TNEF. Conversion is a load followed by a save:

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
    EmailWriteResult result = writer.Write(message, "converted.eml");
    result.RequireNoLoss();
}
```

Use `EmailConversionLossPolicy.Warn` only when the application has made an explicit decision to accept a documented
loss. Opaque MAPI/TNEF metadata that has no portable EML equivalent is reported as a warning while common message
content remains writable. Every artifact writer returns the same `EmailWriteResult`: `SourceSelection` says whether
preserved bytes or regenerated model content was emitted, `AttachmentContentLifetime` records whether bounded content
sources were opened for that operation, `DiagnosticCodes` retains stable evidence, and `LossDisposition` distinguishes
no known loss, explicitly accepted loss, and a blocked write. Transport adapters can carry that result without defining
another message or artifact model.

## Compare message semantics

Semantic comparison hashes a canonical, versioned projection rather than serialized bytes. The migration profile
normalizes store identity and serialization details, the strict profile includes representation details, and the
deduplication profile excludes synchronization and modification state:

```csharp
EmailSemanticComparisonReport comparison = EmailSemanticComparer.Compare(source, destination);

if (!comparison.IsMatch) {
    foreach (EmailSemanticDifference difference in comparison.Differences) {
        Console.WriteLine($"{difference.Kind}: {difference.Path}");
    }
}
```

Difference reports contain canonical paths and lengths, not message values. Before persisting fingerprints for
private mail, provide a random `digestKey` through `EmailSemanticComparisonOptions`; the resulting HMAC-SHA-256
fingerprints cannot be correlated without that caller-owned key.

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

Personal distribution lists use `OutlookItemKind.DistributionList` and `OutlookDistributionList`. Member and One-Off
EntryIDs remain synchronized, validation enforces Outlook's per-property byte limit, and malformed identities keep
their raw bytes and diagnostics. Meeting request/response/cancellation and task request/accept/reject/update items also
have typed lifecycle projections instead of requiring message-class string switches.

Decoded appointment and task recurrence uses one `OutlookRecurrence` model across MSG, TNEF, and PST. Expansion is
always bounded and stays in local-clock space until the caller supplies either a system time zone or the decoded
Outlook time-zone definition:

```csharp
OutlookRecurrenceExpansionResult occurrences = OutlookRecurrenceExpander.Expand(
    item.Appointment!.Recurrence!,
    new OutlookRecurrenceExpansionOptions {
        WindowStart = new DateTime(2026, 7, 1),
        WindowEnd = new DateTime(2026, 8, 1),
        MaxOccurrences = 500
    });

OutlookRecurrenceIcsExportResult exported =
    OutlookRecurrenceIcsConverter.Export(item.Appointment.Recurrence!);
```

Modified and deleted occurrences remain explicit. ICS import/export returns a conversion report and does not silently
drop a recurrence pattern, exception, or time-zone constraint it cannot represent.

An Outlook template uses the same model and compound-file engine as MSG while retaining its template identity:

```csharp
EmailDocument template = EmailDocument.Load("meeting.oft");
template.Subject = "Reusable meeting request";
template.Save("updated.oft");
```

## Standalone iCalendar and vCard documents

`IcsDocument` and `VCardDocument` expose the same ordered content-line model used by MIME projections. Repeated,
grouped, unknown, IANA, and `X-` properties and their parameters remain available for inspection and mutation instead
of being discarded by a narrow typed projection:

```csharp
IcsDocument calendar = IcsDocument.Load("meeting.ics");
ContentLineComponent meeting = calendar.GetComponents("VEVENT").Single();
meeting.SetProperty("SUMMARY", "Updated planning meeting");
meeting.SetTemporalValue("DTSTART",
    IcsTemporalValue.Zoned(new DateTime(2026, 7, 20, 9, 0, 0), "Europe/Warsaw"));

IReadOnlyList<ContentLineValidationIssue> calendarIssues = calendar.Validate();
calendar.Save("updated.ics");

VCardDocument contacts = VCardDocument.Load("contacts.vcf");
ContentLineComponent contact = contacts.Cards[0];
contact.SetVCardText("FN", "Ada Lovelace");
contact.AddProperty("EMAIL", "ada@example.test").SetParameter("TYPE", "work");

IReadOnlyList<ContentLineValidationIssue> contactIssues = contacts.Validate();
contacts.Save("updated.vcf");
```

iCalendar temporal helpers retain `DATE`, floating, UTC, and `TZID`-local forms without resolving identifiers through
the host operating system. RRULE helpers parse and update known recurrence parts while retaining unknown parts.
vCard 2.1, 3.0, and 4.0 syntax is supported, including groups, repeated properties, binary/data-URI values, RFC 6868
parameter escaping, legacy quoted parameters, and quoted-printable continuation. Validation reports conformance
issues; it does not erase vendor data. Legacy `.vcs` vCalendar files use the same preservation model, while validation
continues to report constructs that are not RFC 5545 iCalendar 2.0.

When an appointment or task is written as EML, it becomes a `text/calendar` iCalendar part. Contacts become vCard
attachments. Reminders become `VALARM`; task fields without a direct iCalendar property use valid `X-OFFICEIMO-*`
extensions so they survive an OfficeIMO EML/MSG/TNEF cycle while remaining ignorable to other calendar readers.
Reading those parts restores the corresponding typed model. Source calendar/vCard bytes are retained
while the projected model is unchanged; editing a projected item is blocked by default when regeneration could omit
unmodeled source properties. Meetings whose only attendee data is display text are also blocked because a valid
iCalendar `ATTENDEE` requires an address. Journal and sticky-note models remain lossless across MSG/TNEF, but have no
standard EML representation and therefore require an explicit non-blocking loss policy for EML output.

Properties that do not have a convenience field remain available through the typed `Mapi` property bag. The shared
`MapiKnownProperties` vocabulary covers the standard and Outlook named properties understood by OfficeIMO, while
custom named properties use the same strongly typed key contract:

```csharp
string? displayName = item.Mapi.GetValueOrDefault(MapiKnownProperties.PidTag.DisplayName);

var customName = new MapiPropertyKey<string>(
    "CustomName",
    MapiPropertySets.PublicStrings,
    "CustomName",
    MapiPropertyType.Unicode,
    MapiPropertyType.String8);

string? customValue = item.Mapi.GetValueOrDefault(customName);
item.Mapi.Set(customName, "Retained with the artifact");
```

## RTF bodies

`EmailBody.Rtf` contains the byte-preserving RTF source decoded from `PidTagRtfCompressed`. The MSG and TNEF writers serialize an assigned RTF body with the MS-OXRTFCP compression format, and the reader accepts both the `LZFu` and `MELA` forms.

RTF syntax editing and semantic conversion belong to `OfficeIMO.Rtf`. Generate RTF through that package when the source contains characters that need RTF escapes. `OfficeIMO.Reader` will route an RTF-only email body through the registered `OfficeIMO.Reader.Rtf` handler; without that optional adapter, it preserves the RTF source and reports the fallback explicitly.

## Protected messages

`EmailDocument.Protection` detects opaque and clear-signed S/MIME plus signed or encrypted OpenPGP/MIME wrappers. It
retains the original artifact bytes automatically, including Outlook's complete outer `multipart/signed` attachment
for `IPM.Note.SMIME.MultipartSigned`. `EmailSmime` creates and verifies clear/opaque S/MIME, creates and decrypts
EnvelopedData, and supports sign-then-encrypt through an explicitly supplied security provider; OpenPGP remains outside this package.
Writing an unchanged protected document in its source format emits those bytes verbatim; an edited or cross-format
write is blocked by default because regenerating the wrapper would invalidate its cryptographic meaning.

```csharp
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;
using X509Certificate2 recipient = LoadRecipientCertificate();
using X509Certificate2 signingCertificate = LoadSigningCertificate();
EmailDocument protectedMessage = EmailDocument.Load("message.eml");
IOfficeSecurityProvider security = OfficeSecurityProvider.Default;

EmailSmimeVerificationResult verified = EmailSmime.Verify(protectedMessage, security);
if (verified.IsCryptographicallyValid) {
    Console.WriteLine(verified.SignedContent?.Subject);
}

EmailSmimeDecryptionResult decrypted = EmailSmime.Decrypt(protectedMessage, recipient, security);
if (decrypted.Decrypted) {
    Console.WriteLine(decrypted.DecryptedContent?.Body.Text);
}

EmailSmimeProcessingResult processed = EmailSmime.DecryptThenVerify(
    protectedMessage, recipient, security);
Console.WriteLine(string.Join(", ", processed.ProcessingOrder));

EmailDocument outgoing = new EmailDocument {
    Subject = "Protected report"
};
outgoing.Body.Text = "Attached report details";

EmailSmimeCreationResult signed = EmailSmime.Sign(
    outgoing,
    signingCertificate,
    security,
    EmailSmimeSignatureMode.ClearSigned);
File.WriteAllBytes("signed.eml", signed.Message);

EmailSmimeCreationResult encrypted = EmailSmime.SignAndEncrypt(
    outgoing,
    signingCertificate,
    new[] { recipient },
    security);
File.WriteAllBytes("signed-and-encrypted.eml", encrypted.Message);
```

Certificate/key discovery is intentionally not implicit. Verification accepts caller trust/revocation policy through
`CmsVerificationOptions` and applies the S/MIME email-signing certificate purpose, including the standard
email-protection EKU; creation and decryption require explicitly supplied certificates and a provider. The original
protected document is never mutated, and the decrypted/signed projection is returned separately. Detached verification first
uses the exact retained MIME bytes and, only when necessary, retries the standard CRLF canonical form used by S/MIME
mail clients; `SignedMimeEntity` still exposes the original bytes and a structured diagnostic records the fallback.
Results name the concrete provider and project signer identity, chain-building, revocation, timestamp, and offline-policy
status into stable email diagnostic codes. `DecryptThenVerify` records and enforces the outer-decrypt, inner-verify order,
while an unparseable protected entity remains available as retained bytes plus diagnostics.
The current Bouncy Castle envelope adapter requires an RSA recipient key that the platform certificate handle permits
to be exported. Non-exportable keys fail with `EnvelopePrivateKeyNotExportable` instead of weakening key policy.

## Mailbox stores and store-backed content

PST, OST, Outlook for Mac OLM, and Apple Mail EMLX containers are included in `OfficeIMO.Email`. The Store APIs yield ordinary `EmailDocument` instances while preserving folder paths, store metadata, diagnostics, and bounded attachment behavior.

`EmailStoreSession.IsPstPasswordProtected` and `EmailStoreInspectionReport.IsPstPasswordProtected` expose the PST header's password-protection state. PST uses a checksum-based access deterrent, not cryptographic content encryption; these properties do not claim confidentiality or decrypt content.

```csharp
using OfficeIMO.Email.Store;

using EmailStoreSession session = EmailStoreSession.Open("archive.pst");
EmailStoreItemReference firstReference = session.EnumerateItems(
    new EmailStoreEnumerationOptions(maxItems: 1)).Single();
EmailDocument firstMessage = session.ReadItem(firstReference).Document;
```

Long-running PST/OST/OLM/EMLX/Mbox/mail-directory migrations can opt into an atomic checkpoint. The checkpoint binds
the exact source byte fingerprint, bounded source catalog, destination, conversion options, writer state, item provenance,
and verification journal. Resume rejects changed sources or options instead of silently continuing against different data:

```csharp
var options = new EmailStorePstConversionOptions(
    checkpointPath: "migration.checkpoint",
    checkpointIntervalItems: 500,
    partialResultPolicy: EmailStorePartialResultPolicy.RetainResumableState,
    verifyAfterWrite: true,
    failOnDataLoss: true);

EmailStorePstConversionReport migration = EmailStoreConverter.ConvertToPst(
    "source-mail", "destination.pst", conversionOptions: options);
Console.WriteLine($"{migration.Disposition}; resumed={migration.WasResumed}");
```

Attachment streams remain bounded and operation-scoped. A requested verification manifest records keyed fingerprints,
ordinals, outcomes, and difference-path tokens without subjects, addresses, message content, filenames, or store IDs.
Use `DiscardIncompleteState` when an interrupted migration must not retain its writer-owned staging files.

Use `EmailDataArtifact.Open` from the included `OfficeIMO.Email.Data` namespace when an application wants one discovery entry point across individual artifacts, stores, and OAB files. No additional package is required.

### Store-backed attachment content

`EmailAttachment.Content` remains the simple in-memory representation. A mailbox or archive provider can instead set
`ContentSource` to an `IEmailContentSource` that opens a fresh decoded stream on demand:

```csharp
EmailAttachment attachment = item.Attachments[0];

using Stream content = await attachment.OpenContentStreamAsync(cancellationToken);
await content.CopyToAsync(destination, 81920, cancellationToken);
```

EML, MSG, and TNEF writers consume either representation through the same bounded path. The interface deliberately
contains no PST, OST, MAPI, or Outlook types, so mailbox and store owners can yield ordinary `EmailDocument`
instances while keeping source lifetime and property-stream access in the owning Store API area.

## Safe HTML and text projection

Install the optional `OfficeIMO.Email.Html` bridge when a consumer needs one shared body-selection, RTF-fallback,
sanitization, and resource-resolution contract. It chooses HTML, RTF, or plain text under an explicit policy, blocks
remote resources by default, resolves CID/content-location/filename references, and exposes bounded operation-scoped
resources plus a prepared `HtmlConversionDocument`. `OfficeIMO.Email.Image` and `OfficeIMO.Reader.Email` consume this
same bridge; Reader can then project the prepared safe document to text or Markdown. The core `OfficeIMO.Email` package
does not reference AngleSharp or another HTML engine.

## Resource limits

`EmailReaderOptions` is immutable and applies limits before retaining decoded content. It controls source size, header size and count, MIME part count and depth, per-attachment and aggregate attachment bytes, embedded-message depth, CFB directory entries, MAPI properties and decoded property bytes, and TNEF attributes.

Use `includeAttachmentContent: false` when only attachment metadata is needed. Parsing still validates the source
structure, but ordinary decoded attachment payloads are not retained in the result model. Calendar and vCard parts
that define the typed item are retained because they are semantic message content, not optional file payloads.

## Reader integration

`OfficeIMO.Reader.Email` recognizes `.eml`, `.msg`, `.oft`, `.mbox`, `.mbx`, `.tnef`, `winmail.dat`, `.ics`, `.vcs`,
`.vcf`, `.vcard`, PST, OST, OLM, EMLX, mailbox directories, and OAB data. Calendar and contact files route through the
public `IcsDocument` and `VCardDocument` engines. Store and address-book projections live in that same adapter package;
they add no separate NuGet layer and reuse this package's parsers and models. Rich results include envelope and Outlook
metadata, structured diagnostics, materializable attachment assets, embedded messages, and chunks extracted through
the Reader handlers configured by the host.

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddEmailHandlers()
    .Build();
OfficeDocumentReadResult result = reader.ReadDocument("message.msg");

Console.WriteLine(result.Source.Title);
foreach (OfficeDocumentAsset attachment in result.Assets) {
    Console.WriteLine($"{attachment.FileName}: {attachment.LengthBytes}");
}
```

## Scope boundary

`OfficeIMO.Email` owns offline artifact parsing, serialization, and format-neutral Outlook data. It does not connect to mail servers, authenticate users, send messages, discover certificates or private keys, or implement DKIM, ARC, or OpenPGP. `EmailSmime.Sign`, `Encrypt`, `SignAndEncrypt`, `Verify`, `Decrypt`, and `DecryptThenVerify` are thin data-oriented adapters over a caller-supplied `IOfficeSecurityProvider`: Email owns bounded MIME construction and exact protected-byte handling, while the provider owns CMS, certificates, and trust policy. The adapters create signed or encrypted MIME, verify exact clear-signed bytes and opaque signed-data, decrypt caller-selected EnvelopedData recipients, retain the original protected artifact, and return a separate parsed protected-content document when possible. Certificate/key and provider selection remain explicit and caller-owned.

The package does not expose general-purpose CFB transactions. Its Store API area owns PST, OST, OLM, EMLX, Mbox,
Apple Mail, and Maildir traversal, selection, validation, native export, verified conversion, multi-store merge, and
verified atomic rewrite mutation of an existing Unicode PST. ANSI PST mutation, OST mutation/writing,
in-place NDB editing, append, in-place repair, and password/encryption authoring remain outside the contract. Store
also owns distinct-destination verified compaction and query/size-based split; neither operation edits the open source.

For exact pass-through of an ordinary unprotected artifact, read with `preserveRawSource: true` and write with
`usePreservedRawSource: true`. Protected artifacts use safe unchanged pass-through automatically.

## Dependency footprint

- **External:** no third-party email engine or Outlook interop. `System.Text.Encoding.CodePages` supplies legacy encodings.
- **OfficeIMO:** `OfficeIMO.Core` and `OfficeIMO.Rtf`. MIME, MSG/MAPI, TNEF, mbox, iCalendar, vCard, Store, OAB, protected-wrapper detection, and compressed-RTF handling remain first-party and ship in this package.
- **Optional security:** install `OfficeIMO.Security` for S/MIME CMS/X.509 signing, encryption, verification, and decryption, or provide another `IOfficeSecurityProvider`. It is not a transitive Email dependency.
- **Optional HTML:** install `OfficeIMO.Email.Html` for safe body/resource projection. HTML and RTF-to-HTML dependencies remain outside this core package.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
