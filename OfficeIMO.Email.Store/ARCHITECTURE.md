# OfficeIMO email-store architecture

## Package ownership

The finished package boundary is intentionally small:

| NuGet | Owns |
| --- | --- |
| `OfficeIMO.Email` | `EmailDocument`, MIME/EML, MSG, OFT, TNEF, mbox, MAPI projection, and item writers. |
| `OfficeIMO.Email.Store` | PST/OST/OLM/EMLX and mailbox-directory traversal, sessions, selection, validation, recovery discovery, and export orchestration. |
| `OfficeIMO.Reader.EmailStore` | Optional Reader registration and bounded projection into Reader chunks, metadata, assets, and diagnostics. |

OFT is an individual Outlook template, not a mailbox store. It stays in `OfficeIMO.Email`. The store package may
export a selected store item as OFT, but it does not own the OFT format. MIME primitives also remain inside
`OfficeIMO.Email`; a separate public MIME package is not needed by this design.

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

## OST to PST decision

OST-to-PST output is feasible, but it is a separate major writer project. It is not safe to copy an OST and change
its client signature. A valid PST creator must implement and validate all three format layers:

1. NDB: Unicode header state and checksums, allocation maps, BBT/NBT construction, block/page allocation, BID/NID
   counters, data trees, subnode trees, and an atomic commit strategy.
2. LTP: Heap-on-Node allocation, BTree-on-Heap, property contexts, table contexts, row indexes, row matrices, and
   variable values.
3. Messaging: mandatory store/folder/message properties, hierarchy and contents tables, recipients, attachments,
   embedded messages, named-property maps, search/deleted folders, and entry identifiers.

Microsoft's current PST specification explicitly describes the format as a read/write contract and defines minimum
objects required for a mountable file. It also requires allocation metadata and header state to be maintained. See
the official [MS-PST overview](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-pst/141923d5-15ab-4ef1-a524-6dce75aae546),
[NDB layer requirements](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-pst/9d2083cf-fd37-4a0d-b61a-d2ef10a89a04),
[LTP layer](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-pst/77007716-7993-44fe-9b40-9526157cfc6d),
and [minimum object requirements](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-pst/7af54176-5108-4ac7-973f-8252ad223acb).

A future writer should be append-oriented and Unicode-only, write to a new destination, never mutate the OST, and
remain internal until it passes:

- reopen and deep-validation tests through an independent read path;
- Outlook mount/import interoperability on supported Windows test hosts;
- named-property, typed-item, recipient, attachment, and embedded-message round trips;
- corruption/fault-injection tests around allocation and commit boundaries;
- multi-gigabyte and multi-million-item stress tests with bounded memory;
- conversion manifests that distinguish preserved, normalized, omitted, and server-only data.

Until those gates exist, supported migration outputs are EML/MSG/OFT/TNEF directories and streaming mbox. MSG is
the closest current first-party output when retaining Outlook/MAPI item semantics matters.

## Source and output safety

- Store sessions are read-only and not thread-safe.
- Caller-owned streams stay open and return to their original position when a session is disposed.
- Recovery APIs discover indexed orphans; they do not rewrite source indexes.
- Directory exports use sanitized, stable-ID-suffixed paths and do not overwrite by default.
- Mbox export is streamed to a same-directory temporary file before commit.
- All parsing and writing uses BCL and first-party OfficeIMO code; no native Outlook or third-party email-store parser
  is introduced.
