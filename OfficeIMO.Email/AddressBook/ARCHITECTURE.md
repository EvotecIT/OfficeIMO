# OfficeIMO Email Address Book architecture

## Ownership

OAB and mailbox-store behavior are separate API areas inside one production package:

```text
OfficeIMO.Email                         one NuGet and assembly
├── OfficeIMO.Email.Store               PST/OST/OLM/EMLX mailbox API
└── OfficeIMO.Email.AddressBook         OAB directory snapshot API
    └── OfficeIMO.Reader.Email           optional unified Reader projection
```

An OAB record represents a directory object, not a message. Keeping its model and namespace outside
`OfficeIMO.Email.Store` preserves that semantic boundary without creating another package or dependency edge, and
avoids projecting directory entries as artificial `EmailDocument` instances.

## Read pipeline

1. Bounded discovery classifies `.oab` components and skips reparse points.
2. A session opens each readable v4 Full Details source and reads the 12-byte header, dynamic metadata tables, and schema-defined header record.
3. Reference enumeration reads record sizes only. A reference records the list, entry index, byte offset, and length.
4. `ReadEntry` decodes one record's presence array and property values against that file's schema.
5. Search reuses the same sequential record decoder and checkpoints the exact next offset.
6. Validation reuses the same framing and decoder, optionally adding a streaming CRC pass.

There is one property decoder and one record-framing implementation. Search, validation, random access, Reader, and typed projections do not carry private OAB parsers.

## Large-file behavior

File-backed sources are reopened only for active operations with read/write/delete sharing so Outlook cache replacement is not blocked. Metadata remains resident; records do not. Sequential enumeration and search hold one source stream and one record buffer. Random access opens the selected source and seeks directly to the reference.

Caller-owned streams remain open and return to their original position after each lease. Sessions are not thread-safe because a caller stream can have only one position.

Every retained variable-size structure has a reader limit. Search additionally bounds scanned records, results, searchable characters, and terms. Validation bounds checksum bytes and records. Cancellation is checked during discovery, checksum chunks, and every record loop.

## Compatibility policy

The v4 schema is file-defined. Property counts and optional fields are not hard-coded. Known values receive typed projections; every decoded property remains available as a shared `MapiProperty`, including optional original encoded bytes.

The API reports recognizable v2/v3 and template components without pretending their layouts are v4 records. Unsupported property types or component versions fail explicitly. Compressed Exchange distribution files and differential patch application need a separate, specification-backed decompression/update capability before they can be advertised.

## Security and privacy

OAB content can contain names, addresses, phone numbers, organization data, and distribution-list membership. The library returns values only when a caller requests entries or search results. Discovery and progress models are aggregate. Reader does not emit arbitrary raw properties, and membership values are opt-in.

Tests use generated fixtures. Private Outlook cache validation is aggregate-only and is never copied into the repository.
