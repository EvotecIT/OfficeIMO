# OfficeIMO email-store architecture

## Package ownership

The finished package boundary is intentionally small:

| NuGet | Owns |
| --- | --- |
| `OfficeIMO.Email` | `EmailDocument`, MIME/EML, MSG, OFT, TNEF, mbox, MAPI projection, and item writers. |
| `OfficeIMO.Email.Store` | PST/OST/OLM/EMLX and mailbox-directory traversal, sessions, selection, validation, recovery discovery, and export orchestration. |
| `OfficeIMO.Email.AddressBook` | OAB component discovery, v4 directory entries and distribution lists, bounded search, raw properties, and integrity validation. |
| `OfficeIMO.Reader.EmailStore` | Optional Reader registration and bounded projection into Reader chunks, metadata, assets, and diagnostics. |
| `OfficeIMO.Reader.EmailAddressBook` | Optional typed OAB entry projection into Reader chunks, metadata, and diagnostics. |

OFT is an individual Outlook template, not a mailbox store. It stays in `OfficeIMO.Email`. The store package may
export a selected store item as OFT, but it does not own the OFT format. MIME primitives also remain inside
`OfficeIMO.Email`; a separate public MIME package is not needed by this design.

## Outlook data beyond message stores

OAB data is an offline directory snapshot, not a mailbox container. `OfficeIMO.Email.AddressBook` therefore owns
OAB file-set discovery, bounded entry and distribution-list enumeration, search, raw property retention, integrity
validation, and projection into shared OfficeIMO address/contact models. `OfficeIMO.Reader.EmailAddressBook` is the
corresponding thin Reader adapter; OAB is not an `OfficeIMO.Reader.EmailStore` format registration and an address-book
entry is not disguised as an `EmailDocument`.

The same ownership test applies to other Outlook-local data:

- OFT is already an individual Outlook item and stays in `OfficeIMO.Email`.
- Signatures and stationery are HTML/RTF/text resource sets and belong with an Outlook-profile resource owner if a
  reusable workflow requires them.
- Autocomplete caches, account/profile settings, synchronization state, and search indexes are profile or cache
  artifacts. They need their own evidence, safety limits, and public models before a package claims support.
- Exchange or Microsoft 365 directory synchronization remains a network/provider concern; an offline OAB reader
  must not grow authentication or tenant administration behavior.

This leaves a clean extension seam without making completion of PST/OST/OLM/EMLX support depend on unrelated
profile formats.

## Large-store contract

`EmailStoreSession` is the primary PST/OST API. Opening builds a lightweight folder catalog while NBT entries are
streamed. B-tree pages use a bounded LRU cache. Item enumeration streams contents-table rows and resolves NIDs/BIDs
on demand. `ReadSummary` decodes only browsing properties. `ReadItem` is the explicit boundary that projects one
complete item.

This design keeps memory related to the active parser structures and selected item rather than the store's total
size. It does not promise that one unbounded message or attachment can fit in memory; those operations remain
guarded by the configured per-item limits.

Mailbox-directory sessions index bounded file metadata and open only selected EML/EMLX files. Reparse points are
skipped. OLM is bounded but currently materialized during open because one XML archive entry can contain multiple
logical items; making OLM item payloads lazy would require a durable XML item-location index.

## Managed PST writer and conversion

OST-to-PST conversion never copies a store and changes its client signature. `EmailStorePstWriter` creates a new
Unicode PST through three internal layers:

1. NDB: Unicode header state and checksums, allocation maps, BBT/NBT construction, block/page allocation, BID/NID
   counters, data trees, subnode trees, and an atomic commit strategy.
2. LTP: Heap-on-Node allocation, BTree-on-Heap, property contexts, table contexts, row indexes, row matrices, and
   variable values.
3. Messaging: mandatory store/folder/message properties, hierarchy and contents tables, recipients, attachments,
   embedded messages, named-property maps, search/deleted folders, and entry identifiers.

Microsoft's PST specification describes the format as a read/write contract and defines the minimum objects
required for a mountable file. It also requires allocation metadata and header state to be maintained. See
the official [MS-PST overview](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-pst/141923d5-15ab-4ef1-a524-6dce75aae546),
[NDB layer requirements](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-pst/9d2083cf-fd37-4a0d-b61a-d2ef10a89a04),
[LTP layer](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-pst/77007716-7993-44fe-9b40-9526157cfc6d),
and [minimum object requirements](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-pst/7af54176-5108-4ac7-973f-8252ad223acb).

The public writer is incremental at the item boundary and new-file-only. It allocates data blocks, subnode trees,
BBT/NBT pages, allocation maps, property contexts, table contexts, and the mandatory store and folder objects,
then commits a same-directory temporary file. `EmailStoreSession.ExportToPst` remains a thin orchestration surface:
it enumerates the source, maps its folder tree, streams selected attachments, and reports source items that cannot
be represented. All message and typed-item property projection stays in `OfficeIMO.Email` and is shared with MSG.

Current validation covers:

- reopen and structural CRC/signature validation through the OfficeIMO reader;
- multi-block heaps, multi-leaf table indexes, large attachment data trees, embedded messages, associated items,
  recipients, named properties, and fixed and variable multi-valued properties;
- synthetic OST-to-new-PST conversion with a byte-for-byte unchanged source;
- independent libpff open and synthetic message export;
- an opt-in classic Outlook mount/read/remove test for Windows interoperability hosts;
- an opt-in private-corpus lane that reports only aggregate structure and diagnostic counts.

This is a projection conversion, not an Exchange recovery service. Search results are materialized as static
folders; a source Name-to-ID entry that is unavailable receives a diagnostic placeholder mapping; and attachment
or OST content absent from the local source is reported rather than fabricated. The writer does not append to,
repair, compact, password-protect, encrypt, or otherwise mutate existing PST/OST files. ANSI PST and OST output are
also outside the current contract.

## Source and output safety

- Store sessions are read-only and not thread-safe.
- Caller-owned streams stay open and return to their original position when a session is disposed.
- Recovery APIs discover indexed orphans; they do not rewrite source indexes.
- Directory exports use sanitized, stable-ID-suffixed paths and do not overwrite by default.
- Mbox export is streamed to a same-directory temporary file before commit.
- All parsing and writing uses BCL and first-party OfficeIMO code; no native Outlook or third-party email-store parser
  is introduced.
