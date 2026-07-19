# OfficeIMO.Email.AddressBook

`OfficeIMO.Email.AddressBook` is the Outlook Offline Address Book API area included in the `OfficeIMO.Email` package. It reads component sets without Outlook, native code, or a third-party OAB parser, inventories legacy components, and lazily decodes uncompressed OAB v4 Full Details files into shared address, contact, and MAPI models.

The API is read-only. It does not modify an Outlook cache or contact Exchange.

## Install

```powershell
dotnet add package OfficeIMO.Email
```

## Inspect an OAB cache

Inspection reports component roles without requiring a readable Full Details file:

```csharp
using OfficeIMO.Email.AddressBook;

OfflineAddressBookDiscoveryReport inventory =
    OfflineAddressBookInspector.Inspect(@"C:\Users\me\AppData\Local\Microsoft\Outlook\Offline Address Books");

foreach (OfflineAddressBookFileInfo file in inventory.Files) {
    Console.WriteLine($"{file.Format}: {file.Length:N0} bytes; entries={file.CanEnumerateEntries}");
}
```

Directory discovery is deterministic, bounded by file count and depth, and does not follow reparse points. It identifies v4 Full Details files, display templates, and the standard v2/v3 Browse, ANR, RDN, Details, and Changes roles.

## Enumerate entries lazily

Opening a directory creates one list descriptor per readable Full Details component. Records are opened and decoded as the caller enumerates them:

```csharp
using OfflineAddressBookSession session = OfflineAddressBookSession.Open("Offline Address Books");

foreach (OfflineAddressBookEntryReference reference in session.EnumerateEntryReferences(
    new OfflineAddressBookEnumerationOptions(maxEntries: 1_000))) {
    OfflineAddressBookEntry entry = session.ReadEntry(reference);
    Console.WriteLine($"{entry.DisplayName}: {entry.SmtpAddress}");
}
```

`OfflineAddressBookEntry` exposes names, SMTP/X500/proxy addresses, organization and postal fields, phones, object/display types, distribution-list membership, truncation markers, and the complete schema-defined `MapiProperty` collection. Original encoded property bytes can be retained for fidelity analysis. Compatible entries project through `ToEmailAddress()`, `ToOutlookContact()`, and `ToSummary()` instead of introducing duplicate email/contact models.

Caller-owned streams must be readable and seekable. The session uses the stream's current position as the component start, leaves the stream open, and restores its original position after every operation.

## Search large address books

Search is a bounded offline scan. Exact record offsets make checkpoints inexpensive to resume even late in a large file:

```csharp
OfflineAddressBookSearchCheckpoint? checkpoint = null;
do {
    OfflineAddressBookSearchReport batch = session.Search(
        new OfflineAddressBookSearchQuery(
            new[] { "engineering", "example.test" },
            fields: OfflineAddressBookSearchFields.Organization |
                    OfflineAddressBookSearchFields.Addresses,
            matchMode: OfflineAddressBookSearchMatchMode.AllTerms,
            maxEntriesScanned: 50_000,
            maxResults: 250,
            resumeFrom: checkpoint));

    foreach (OfflineAddressBookSearchResult match in batch.Results) {
        Console.WriteLine($"{match.Summary.DisplayName}: {match.Snippet}");
    }
    checkpoint = batch.NextCheckpoint;
} while (checkpoint != null);
```

Queries can filter one address list or projected object type and search names, addresses, organization, phones, postal fields, comments, and membership. Term count, term length, per-entry searchable characters, records scanned, matches, progress interval, and error handling are explicit.

Checkpoints belong to the same session snapshot and query scope. Reopen the session if Outlook replaces a source component.

## Resolve Outlook and Exchange identities offline

Build one bounded immutable index when PST/MSG recipients contain legacy EX/X.500 identities or aliases instead of a
portable SMTP address:

```csharp
OfflineAddressBookIdentityIndex identities = session.BuildIdentityIndex(
    new OfflineAddressBookIdentityIndexOptions(
        maxEntries: 1_000_000,
        maxIdentitiesPerEntry: 64));

OfflineAddressBookIdentityResolution resolution =
    identities.Resolve("/o=Example/ou=Recipients/cn=ada", "EX");

if (resolution.Status == OfflineAddressBookIdentityResolutionStatus.Resolved) {
    Console.WriteLine(resolution.Candidate!.PrimarySmtpAddress);
}
```

The index covers primary SMTP, proxy, EX/X.500, target, and optional account/display-name values. It retains every
candidate for a duplicate key and reports `Resolved`, `Ambiguous`, `NotFound`, or `Incomplete`; it never silently selects the first
directory entry. Entry and per-entry identity limits are part of the result's completeness diagnostics.

## Validate integrity

Validation uses the OAB header's seeded IEEE CRC directly and can stop after the checksum, walk record framing, or decode every selected record against the file-defined schema:

```csharp
OfflineAddressBookValidationReport validation = session.Validate(
    new OfflineAddressBookValidationOptions(
        mode: OfflineAddressBookValidationMode.FullDecode,
        maxEntriesPerAddressList: 1_000_000));

Console.WriteLine($"Valid: {validation.IsValid}; records: {validation.RecordsScanned:N0}");
```

Checksum bytes, records, progress, cancellation, and per-record failures remain bounded and visible. Validation never repairs or rewrites a component.

## Format support

| Component | Contract |
| --- | --- |
| Uncompressed OAB v4 Full Details (`0x00000020`) | Schema/header inspection, lazy entry and distribution-list enumeration, raw properties, search, CRC/framing/full-decode validation |
| OAB display template (`0x00000007`) | Identified by inspection; not decoded as address entries |
| OAB v2/v3 Browse, ANR, RDN, Details, and Changes files | File-set role inspection only; entry decoding is explicitly unsupported |
| Exchange OAB manifests, compressed LZX downloads, and differential patches | Outside the current Outlook-expanded cache reader |

The v4 parser follows Microsoft's [MS-OXOAB specification](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxoab/b4750386-66ec-4e69-abb6-208dd131c7de), including dynamic property tables, presence arrays, compact integers, UTF-8 OAB Unicode values, String8 code pages, multi-valued properties, truncation markers, and the header CRC contract.

## Limits and large files

The default profile permits one component up to 64 GiB while bounding discovered files, directory depth, metadata, schema size, record size, strings, binary values, multi-value counts, and declared entries. Narrow these values for untrusted inputs:

```csharp
var options = new OfflineAddressBookReaderOptions(
    maxInputBytes: 8L * 1024 * 1024 * 1024,
    maxRecordBytes: 8 * 1024 * 1024,
    maxValuesPerProperty: 20_000,
    retainRawPropertyBytes: false);
```

Memory tracks metadata and the active record, not total file size. Enumeration, search, and validation are synchronous streaming APIs with cancellation and aggregate progress.

## Boundaries and dependencies

- `OfficeIMO.Email.AddressBook` owns OAB file-set discovery, v4 schema/record decoding, search, offline identity resolution, and validation.
- `OfficeIMO.Email` owns reusable `EmailAddress`, `OutlookContact`, `MapiProperty`, and diagnostics models.
- `OfficeIMO.Reader.Email` owns optional Reader chunks and registration for OAB alongside other email data.
- The `OfficeIMO.Email.Store` API remains the PST/OST/OLM/EMLX mailbox owner; an OAB entry is not an `EmailDocument`.
- Exchange/Graph directory synchronization, Outlook profile settings, autocomplete caches, search indexes, and address-book mutation are separate concerns.

Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.

The implementation ships in `OfficeIMO.Email`. External parser dependencies: none.
