# Access support and feasibility

The public contract covers a typed model, bounded native inspection and portable catalog/schema/row reading for qualified unprotected Jet 4 and ACE layouts. [CAPABILITIES.md](CAPABILITIES.md) and [operations.json](operations.json) are generated from `AccessCapabilities.cs`. The executable consumer in `OfficeIMO.Access.Verification` owns the public example contract. Native persistence and application-payload decoding remain unsupported.

## Qualified foundation

The library builds for `netstandard2.0`, `net8.0`, `net10.0`, and Windows `net472`. The synthetic Jet 4/ACE 12 corpus comes from the installed Microsoft DAO 16 engine, independently reopens read-only, and exports typed schema and values. The manifest records byte lengths, SHA-256, producer, profile, fixture license and limitations. `OfficeIMO.Access.Tests/Fixtures` contains only synthetic project-owned data.

Foundation tests cover stream ownership and position, byte/page budgets, malformed and unknown headers, cancellation, rollback identity, detached objects, omitted versus null values, binary ownership, reader leases, typed relationships, source replacement and stale reports. They also prove that unsupported saves leave nonexistent destinations and existing streams untouched. A separate public consumer exercises these contracts without internal access.

## Profiles and qualification boundaries

| Profile | Header code / page size | Independent producer evidence | Current operations |
| --- | --- | --- | --- |
| Jet 4 MDB | `1` / 4096 | DAO 16 tables, typed values, indexes, relationships and queries; Access format 9 (2000) and 10 (2002/2003) files | Catalog, schema, properties, rows, index definitions, relationships and query records |
| ACE 12 ACCDB | `2` / 4096 | DAO 16 scalar/fragmented tables, attachments and multivalued fields, read-only schema/value exports | Same read API, plus qualified structured values |
| Jet 3 MDB | `0` / 2048 | Representative Access 97 producer fixture still required | Header recognition only |
| ACE 14 | `3`, subversion `1` / 4096 | DAO calculated field and independent value observation; separate password-required fixture | Unprotected catalog/rows; calculated payload is exact opaque data with diagnostics. Protected catalogs remain unavailable |
| ACE 16 | `5` / 4096 | DAO BigInt definition and independently observed `5,000,000,000` value | Catalog/rows and native signed 64-bit integers |
| ACE 17 | `6` / 4096 | DAO type `26` values outside ordinary Date/Time range and controlled fractional probe independently read by DAO | Catalog/rows, Large Number and Date/Time Extended with scale-sensitive 100 ns precision |

Generation codes are physical compatibility gates, not a claim about the producing application's marketing version. Modern Access produced the format 9/10 fixtures; this is not execution evidence from original Access 2000/2003 applications. Unknown signatures/generations are rejected before a page size is assumed. Header-only inspection keeps catalogs not decoded, protection not assessed and structure unvalidated beyond alignment. Loading an unprotected supported profile decodes the catalog separately. Path loading rejects cross-family suffixes and compiled/add-in/package/template suffixes; stream loading uses content. No reader operation authenticates users or decrypts protected pages.

## Native reader contract

The A02/A03 reader uses one schema and `DbDataReader` API for MDB and ACCDB. Required system catalogs are separate from selected user tables. The catalog retains unknown object types and exact native records; form/report/macro/VBA payload collections remain `NotDecoded` even when tables are decoded. Linked definitions expose redacted connection metadata and never open their targets. Explicit raw catalog records may contain credentials and require caller-controlled handling.

| Feature | Qualified read behavior and evidence |
| --- | --- |
| Catalog and properties | Table/column names, storage flags, database code page/sort metadata and MR2 property maps; unknown maps retain exact bytes and diagnostics |
| Rows | Forward-only owned-page scans with deleted rows skipped, overflow pointers followed, fragmented/grown row layout and declared live-count verification at EOF |
| Indexes and relationships | Primary/unique/foreign flags, ordered fields, descending members and composite cascade relationships; index-root page validation. Index key lookup and constraint execution are outside this reader contract |
| Numeric values | Boolean, Byte, Int16, Int32/AutoNumber, Single, Double, Currency, qualified Decimal precision up to 28 and Large Number Int64. Wider/undeclared Decimal precision retains opaque bytes |
| Text and binary | Jet 4/ACE UTF-16 and compressed Unicode, Polish/Arabic/CJK/emoji, null versus empty, Memo/long text, inline/chained OLE and binary bytes. Database code-page metadata is retained; Jet 3 legacy byte encodings remain unqualified |
| Dates and identities | Ordinary Date/Time including pre-epoch/leap dates, GUID and extended Date/Time with persisted scale 0–7; dates have unspecified timezone |
| Field presentation | Rich-text markup, hyperlink strings and lookup properties/keys retained without rendering, navigation or value substitution |
| Complex values | Multivalued text and attachments retain parent key, typed backing schema/rows and exact encoded content; attachment decode checks the native envelope, zlib checksum and expansion limit. Other backing structures expose native rows without claiming a qualified convenience interpretation |
| Queries | Exact native records and parameter inventory; qualified simple single-table SELECT (including wildcard/expressions) and two-part UNION reconstruction. Sized parameters, external-source qualifiers and other unqualified shapes keep `HasSql=false` and diagnostics; no SQL is executed |
| Calculated/unknown fields | Exact `AccessOpaqueValue`, expression metadata and column diagnostics; no evaluation or silent coercion |
| Security metadata | Catalog owner SID bytes and system permission records are inspectable; no authentication, permission enforcement or security-policy changes |

Loading retains one bounded file snapshot, rather than mapping the file or materializing all rows. Defaults are 64 MiB input, 16,384 pages, 4,096 catalog objects, 4 MiB total decoded metadata, 32 MiB per value/attachment, one million rows traversed per reader and 16,384 chain links. Total metadata accounting includes repeated references. Table selection limits user-schema decoding; required system metadata still consumes its budget. Row/field decoding is lazy. `IsDBNull` does not load a long value, and binary streams advance through native chunks. Attachment content and ordinary `GetValue` binary access allocate only when requested. Complex scans count every backing row traversed against their row limit.

Readers expose deterministic schema before the first row and support `DataTable.Load`. Cancellation and document disposal invalidate traversal/streams. Native documents are immutable, including loads requested with read/write access. Malformed definition/overflow/long-value chains, bad page/row references, truncation, duplicate schema names and excessive budgets fail explicitly. A damaged late user row does not prevent opening its earlier valid rows; failure occurs when the scan reaches it. This bounds the qualification claim to traversed/decoded structures, rather than certifying every page on load.

## MDB and ACCDB feature assessment

`AssessSave` records the following named target losses before any future codec can write. All native save/conversion operations remain unsupported; these assessments never authorize automatic flattening.

| Feature/target | Diagnostic code |
| --- | --- |
| Attachments, multivalued and other complex fields to MDB | `access.conversion.loss.complex.<kind>` |
| Rich text to MDB | `access.conversion.loss.rich-text` |
| Large Number to Jet or ACE 12/14 | `access.conversion.loss.large-number` |
| Date/Time Extended to profiles other than ACE 17 | `access.conversion.loss.extended-date` |
| Calculated field to Jet or ACE 12 | `access.conversion.loss.calculated` |
| Opaque property metadata to a different family/profile without a qualified mapping | `access.conversion.loss.opaque-properties` |

Unsupported property maps, calculated values and application carriers stay exact opaque metadata or explicitly undecoded. Opaque property mappings are diagnosed when the target family/profile changes. The unavailable writer blocks persistence and preservation claims for these definitions.

## Reusable owners

| Capability | Owner and boundary |
| --- | --- |
| Document access/persistence and loss/conflict vocabulary | OfficeIMO.Core; Access uses the established public contracts |
| Bounded snapshots and caller-stream ownership | OfficeIMO.Core's stream reader |
| Physical catalog, pages, allocation, indexes, rows and application carriers | OfficeIMO.Access; ZIP or compound-file APIs cannot substitute for an Access codec |
| MS-OVBA decompression, VBA metadata and signature canonicalization | Existing OfficeIMO.Core primitives are reusable after an Access-specific carrier yields qualified project data |
| VBA signature packaging | Existing OPC carrier handling is not proof for Access's direct project or ACCDC package carriers |
| Preservation and loss assessment | Shared operation vocabulary; Access must supply generation-specific semantic and byte-preservation proof |
| Report drawing and conversion destinations | OfficeIMO.Core Drawing and the owning document engines after Access report decoding/evaluation is qualified |
| SQL execution, providers and database transactions | DbaClientX; it has no Access provider in the inspected baseline. No database execution engine is introduced here |

## Independent corpus and native feasibility

`New-AccessCorpus.ps1` creates fresh databases through DAO, with indexed tables, a relationship, typed rows and a parameterized query. `New-AccessApplicationCorpus.ps1` uses a new owned Access instance with `AutomationSecurity=3`, creates only new synthetic databases, exports forms/reports/module definitions and PDF output, and closes the instance. It does not change trust policy or invoke module code. Its module import uses the documented but unsupported `LoadFromText` oracle surface; native carrier qualification remains separate. Producer build evidence is Microsoft Access `16.0.20430.20092`, 64-bit, and DAO `16.0`.

The table corpus is checked in with independent expected-value exports. Schema, index definitions and relationships are observed after a read-only DAO reopen. Application-object fixtures and exports are also project-owned synthetic evidence. Access creates and re-exports a form, report, inert VBA module and inert `StopMacro` action macro in both file families. Macro text import uses the `SaveAsText` envelope rather than clipboard XML. No macro or module is executed.

`New-AccessProfileCorpus.ps1` creates the profile/protection corpus under `Fixtures/Profiles`. DAO rejects an open without the synthetic password and independently reopens the protected files with it. OfficeIMO still reports protection as not assessed and does not decrypt these files. Modern feature definitions raise the physical header compatibility gates independently of the producer's application version. Jet 3 creation through the installed engine fails with “Could not find installable ISAM”; a genuine legacy producer remains required. Signature carriers and additional protection variants remain qualification work in A04/A09.

`New-AccessReaderCorpus.ps1` produces [the scalar/complex corpus](../OfficeIMO.Access.Tests/Fixtures/Readers/manifest.json), including deleted/grown rows, composite indexes/relationships, Unicode/Memo/OLE, exact Currency/Decimal boundaries, rich text, lookup properties, multivalued text and attachment files. The installed ACE OLE DB engine defines Decimal precision/scale in this validation-only producer; DAO independently reopens the files read-only and exports schema, rows, complex children and decoded attachment bytes. Neither engine is a product dependency.

`New-AccessGenerationCorpus.ps1` produces [the generation corpus](../OfficeIMO.Access.Tests/Fixtures/Generations/manifest.json). Access explicitly creates format 9 and 10 MDB files; independent read-only DAO observations retain their values and saved query text. Synthetic self/password links are inventory-only. Separate DAO-produced ACE files cover calculated representations, Int64 and extended dates. The installed producer persists the extended date at scale zero; a controlled 42-byte mutation creates the seven-digit fractional probe, and an independent read-only DAO consumer confirms all fractional digits. The manifest distinguishes this controlled mutation from an Access-produced value. Native reader tests compare the persisted scale and do not silently round it.

`Test-NativeBootstrap.ps1` is a negative control: it writes a header and page-type skeleton without copying a seed file. DAO rejects both MDB and ACCDB controls (`0x800A0C0F`). This proves that header recognition is insufficient; it does **not** prove template-free native creation. No native writer is enabled by this result.

`NativeBootstrapProbe.cs` and `Test-NativeCreation.ps1` qualify `native-bootstrap-01` and `native-table-01` for a fixed unprotected Jet 4/ACE 12 schema. The managed generator reads no seed or template. It emits the masked header, empty user slots, global/table/long-value allocation maps, system definitions/catalog/permissions rows, index leaves, two user tables, typed rows and a relationship from logical definitions. DAO independently opens both generated files, exports their persisted schema/index/relationship/value definitions, and seeks the primary index. On separate owned verification copies, it rejects duplicate primary keys and invalid foreign keys. [The native manifest](../OfficeIMO.Access.Tests/Fixtures/Native/manifest.json) records the exact bytes and observations.

This establishes feasibility, not a general production writer. The spike has a fixed schema and creation date, a restricted observed ASCII General legacy collation, an observed unprotected header field at `0x6A` whose wider semantics are unqualified, and no modern complex system catalog. Header-derived SID masking must remain consistent with the creation date and password region; the probe grants built-in permissions only inside freshly generated files. Production encoding, general schemas/values, structural edits and additional security profiles remain A05/A06/A09 work. `AccessDocument.Save` continues to fail before output.

The independent application corpus exposes different storage boundaries: Jet 4 has `MSysAccessObjects` with a storage-specific `Data` type; ACE has `MSysAccessStorage` with hierarchy/type/date metadata and an `Lv` payload. MDB Tools 1.0.0 identifies Jet's `Data` field as an unknown physical type (`0x11`), demonstrating a comparison-reader gap. Creation and text/PDF export through Access do not prove OfficeIMO carrier decoding. The Access adapters must extract and preserve form/report/macro/VBA storage before passing qualified project content to shared Core VBA/security primitives. Direct VBA signature and ACCDC distribution carriers remain separately unqualified. A04/A07 own these native carrier read/write criteria.

## Engine alternatives and deployment

The selected product boundary remains managed native codecs with no new external runtime dependency. The installed Windows Access/DAO engine is an independent qualification tool. Adopting ACE/DAO as an optional product provider would require a separate DbaClientX integration, explicit dependency approval, Windows/bitness/install policy and documented deployment limits. It would not satisfy template-free native creation.

The independently maintained [Jackcess](https://github.com/jahlborn/jackcess) Java library has an Apache-2.0 license and useful codec coverage; its database creation uses bundled empty-database resources, so adopting that creation path would not satisfy this program's seed-free requirement. [MDB Tools](https://github.com/mdbtools/mdbtools) documents Jet 3/4 page structures and provides independent reading/export; its libraries and tools have different LGPL/GPL boundaries. Neither engine nor its code/assets is referenced, vendored or shipped by OfficeIMO.Access. Reading format documentation does not approve a production dependency.

Primary references: [DAO database creation](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/dbengine-createdatabase-method-dao), [Access automation security](https://learn.microsoft.com/en-us/office/vba/api/access.application.automationsecurity), [Access object text import](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/application-load-from-text), [Access VBA and package signatures](https://support.microsoft.com/en-us/access/show-trust-by-adding-a-digital-signature-to-an-access-database), and [MDB Tools physical-format notes](https://github.com/mdbtools/mdbtools/blob/dev/HACKING.md).
