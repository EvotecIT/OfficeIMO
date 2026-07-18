# OfficeIMO.Reader.EmailAddressBook

`OfficeIMO.Reader.EmailAddressBook` is a thin Outlook `.oab` Full Details adapter package for `OfficeIMO.Reader`. Parsing stays in the `OfficeIMO.Email.AddressBook` API; this adapter only selects entries and projects typed directory fields into deterministic Reader chunks and document results.

## Install and register

```powershell
dotnet add package OfficeIMO.Reader.EmailAddressBook
```

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.EmailAddressBook;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddEmailAddressBookHandler(new ReaderEmailAddressBookOptions {
        MaxEntries = 1_000
    })
    .Build();

OfficeDocumentReadResult result = reader.ReadDocument("udetails.oab");
```

One entry produces one `ReaderInputKind.Email` chunk with stable list/record identifiers and logical paths. The projection includes typed names, addresses, organization, phones, postal fields, comments, proxy addresses, distribution-list counts, and truncation counts. It never dumps arbitrary raw MAPI properties.

## Stream entries from a large cache

Use the item-at-a-time API for a directory containing multiple lists or whenever an aggregate document would be too large:

```csharp
var options = new ReaderEmailAddressBookOptions {
    MaxEntries = 10_000,
    Query = new OfficeIMO.Email.AddressBook.OfflineAddressBookSearchQuery(
        new[] { "example.test", "engineering" },
        fields: OfficeIMO.Email.AddressBook.OfflineAddressBookSearchFields.Addresses |
                OfficeIMO.Email.AddressBook.OfflineAddressBookSearchFields.Organization,
        maxEntriesScanned: 500_000,
        maxResults: 10_000)
};

foreach (ReaderEmailAddressBookEntryResult entry in reader.ReadEmailAddressBookEntries(
    "Offline Address Books",
    new ReaderOptions { ComputeHashes = false },
    options)) {
    foreach (ReaderChunk chunk in entry.Chunks) {
        Console.WriteLine($"{chunk.Id}: {chunk.Text.Length} characters");
    }
}
```

The query is evaluated by the core OAB session and returns exact references. Reader then projects only those entries. `MaxEntries`, core record/search limits, `ReaderOptions.MaxInputBytes`, `ReaderOptions.MaxChars`, cancellation, and per-entry diagnostics remain effective.

Distribution-list member distinguished names are disabled by default because they can be large and sensitive. Set `IncludeMembershipValues = true` and choose `MaxMultiValueItems` when the host explicitly needs them.

Complete source hashing is also disabled by default so selective ingestion does not force another full pass over a large OAB. Chunk hashes still follow `ReaderOptions.ComputeHashes`; set `ComputeSourceHash = true` only when the host accepts the cost.

## Boundaries

- `OfficeIMO.Email.AddressBook` owns OAB discovery, schema decoding, search, validation, and typed entries.
- `OfficeIMO.Reader.EmailAddressBook` owns registration, selection bounds, safe chunk text, logical paths, hashes, and Reader diagnostics.
- The adapter reads uncompressed v4 Full Details files. Legacy/template components remain inspection-only in `OfficeIMO.Email`.

Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.

The adapter depends on `OfficeIMO.Reader` and the unified `OfficeIMO.Email` package. No parser, native runtime, Outlook automation, or third-party dependency is added by this adapter.
