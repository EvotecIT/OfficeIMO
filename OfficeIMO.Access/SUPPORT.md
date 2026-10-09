# Access support

The public contract covers a typed model, bounded native reading and application inspection, qualified Jet 4/ACE 12/14 VBA editing, unchanged snapshot preservation, and seed-free Jet 4/ACE 12 table creation. [CAPABILITIES.md](CAPABILITIES.md) and [operations.json](operations.json) are generated from `AccessCapabilities.cs`. The executable consumer in `OfficeIMO.Access.Verification` owns the public examples. General native row/schema editing and conversion remain unqualified.

## Qualified foundation

The library builds for `netstandard2.0`, `net8.0`, `net10.0`, and Windows `net472`. The synthetic Jet 4/ACE 12 corpus comes from the installed Microsoft DAO 16 engine, independently reopens read-only, and exports typed schema and values. The manifest records byte lengths, SHA-256, producer, profile, fixture license and limitations. `OfficeIMO.Access.Tests/Fixtures` contains only synthetic project-owned data.

Foundation tests cover stream ownership and position, byte/page budgets, malformed and unknown headers, cancellation, rollback identity, detached objects, omitted versus null values, binary ownership, reader leases, typed relationships, source replacement and stale reports. They also prove that unsupported saves leave nonexistent destinations and existing streams untouched. A separate public consumer exercises these contracts without internal access.

## Profiles and qualification boundaries

| Profile | Header code / page size | Independent producer evidence | Current operations |
| --- | --- | --- | --- |
| Jet 4 MDB | `1` / 4096 | DAO 16 tables, typed values, indexes, relationships and queries; Access format 9 (2000) and 10 (2002/2003) files; independent creation reopen/edit/save | Catalog/schema/rows, application/VBA inventory, whole-file preservation and qualified creation; compact designers are preserve-only |
| ACE 12 ACCDB | `2` / 4096 | DAO 16 scalar/fragmented tables, attachments and multivalued fields; independent application exports and creation reopen/edit/save | Same read/preserve/create API, structured-value reading and qualified version-21 designer metadata |
| Jet 3 MDB | `0` / 2048 | Representative Access 97 producer fixture still required | Header recognition only |
| ACE 14 | `3`, subversion `1` / 4096 | DAO calculated field and independent value observation; native rich designer and code-behind editing; separate password-required fixture | Unprotected catalog/rows and qualified VBA persistence; calculated payload is exact opaque data with diagnostics. Protected catalogs remain unavailable |
| ACE 16 | `5` / 4096 | DAO BigInt definition and independently observed `5,000,000,000` value | Catalog/rows and native signed 64-bit integers |
| ACE 17 | `6` / 4096 | DAO type `26` values outside ordinary Date/Time range and controlled fractional probe independently read by DAO | Catalog/rows, Large Number and Date/Time Extended with scale-sensitive 100 ns precision |

Generation codes are physical compatibility gates, not a claim about the producing application's marketing version. Modern Access produced the format 9/10 fixtures; this is not execution evidence from original Access 2000/2003 applications. Unknown signatures/generations are rejected before a page size is assumed. Header-only inspection keeps catalogs not decoded, protection not assessed and structure unvalidated beyond alignment. Loading an unprotected supported profile decodes the catalog separately. Path loading rejects cross-family suffixes and compiled/add-in/package/template suffixes; stream loading uses content. No reader operation authenticates users or decrypts protected pages.

## Native reader contract

The reader uses one schema and `DbDataReader` API for MDB and ACCDB. Required system catalogs are separate from selected user tables. The catalog retains unknown object types and exact native records; application inventory and per-object decoding have separate availability. Linked definitions expose redacted connection metadata and never open their targets. Explicit raw catalog records may contain credentials and require caller-controlled handling.

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

Readers expose deterministic schema before the first row and support `DataTable.Load`. Cancellation and document disposal invalidate traversal/streams. Native table/schema mutation is unqualified; application VBA uses the explicit persistence workflow below. Malformed definition/overflow/long-value chains, bad page/row references, truncation, duplicate schema names and excessive budgets fail explicitly. A damaged late user row does not prevent opening its earlier valid rows; failure occurs when the scan reaches it. This bounds the qualification claim to traversed/decoded structures, rather than certifying every page on load.

## MDB and ACCDB feature assessment

`AssessSave` records the following named target losses. Cross-family/profile conversion remains unsupported; these assessments never authorize automatic flattening.

| Feature/target | Diagnostic code |
| --- | --- |
| Attachments, multivalued and other complex fields to MDB | `access.conversion.loss.complex.<kind>` |
| Rich text to MDB | `access.conversion.loss.rich-text` |
| Large Number to Jet or ACE 12/14 | `access.conversion.loss.large-number` |
| Date/Time Extended to profiles other than ACE 17 | `access.conversion.loss.extended-date` |
| Calculated field to Jet or ACE 12 | `access.conversion.loss.calculated` |
| Opaque property metadata to a different family/profile without a qualified mapping | `access.conversion.loss.opaque-properties` |

Unsupported property maps, calculated values and application carriers stay exact opaque metadata or explicitly undecoded. Opaque property mappings are diagnosed when the target family/profile changes. Same-profile whole-file preservation retains these definitions without interpreting or rebuilding them.

## Application and VBA inspection

The Access adapter reads Jet `MSysAccessObjects` compound storage and ACE `MSysAccessStorage` hierarchies. It retains native row/catalog identities, storage paths, base/delta/compiled streams and exact opaque payloads. Malformed hierarchy cycles, ambiguous paths and excessive metadata budgets fail explicitly. Ambiguous object-directory slots leave their application group preserve-only. Directory, designer and macro parsing debit the aggregate metadata allowance before copying input, including unsupported representations. `DecodeApplicationObjects = false` leaves this inventory `NotDecoded` while preserving the full snapshot.

| Content | Qualified interpretation | Preserve-only boundary |
| --- | --- | --- |
| Forms and reports | ACE designer version 21 property trees, native sections/control kinds, Name/Caption, record/control/row sources, explicit size properties and Click bindings | Jet compact version 19, unknown node/property semantics and delta merging; missing defaults are not invented |
| Action macros | Independently observed 76-byte single `StopMacro` definition, standalone and embedded | Other actions/argument layouts retain exact streams; action macros are distinct from VBA |
| Table data macros | MR2 wide-property map, qualified Access XML namespace, event/name and top-level statement names; exact XML retained with DTD/resolver disabled | No action execution, expression evaluation or editing |
| Resources | `MSysResources` identity/type/name metadata and attachment access through the existing reader | No automatic theme unpacking, rendering or OLE activation |
| Startup settings | Qualified database property maps, including AppTitle and stored AccessVersion | Unknown property records retain exact payloads |
| VBA | Shared Core MS-OVBA directory/source inspection, declared modules, source offsets/flags and stored library references; CP1250 source fixture includes Polish text | Missing compiled-source text, unsupported code pages, built-in implicit references and signatures; referenced libraries are never loaded |
| Dependencies | Observed relationship fields, qualified query table records and designer sources, with unresolved references labeled | Partial inventory; no arbitrary SQL/VBA parsing, expression resolution or automatic rewrites |

`ChangeJournal` records object identities, revisions and operations; rollback removes its entries. Qualified native VBA application records `vba.apply`; native row and schema edits remain unqualified.

VBA inspection shares one expansion allowance across the directory and module sources, bounded by `MaxMetadataBytes`. Bytes expanded before malformed source fails still consume that allowance. The native storage reader bounds encoded input before supplying project streams. Duplicate module names or stream identities make the project inventory opaque, so the inspector does not decode the same stream repeatedly. After expansion-budget exhaustion, declared module metadata remains available and unread source remains opaque with a limit diagnostic. Cancellation reaches module and compressed-chunk traversal.

## Native VBA persistence

`GetVbaProject` detaches a bounded compound project from native Jet storage or the ACE storage hierarchy. `SetVbaProject` builds and rereads a complete native candidate before staging it. Core owns source normalization, PROJECT declarations, MS-OVBA directory/compression and compilation-cache invalidation; Access owns native carriers, ordinary-module catalog/directory slots, permissions, allocation and index updates.

| Area | Qualified boundary |
| --- | --- |
| Input | Unprotected Jet 4 and ACE 12/14 applications with decoded source and qualified storage, permission and index layouts |
| Modules | Source replacement; standard/class add, remove and rename; first project in an existing native empty application |
| Code-behind | Existing native form/report `DocClass` modules with their original name, kind and `VB_Base`; designer streams and event bindings remain exact |
| Identity | Existing ordinary module native IDs, storage slots and logical catalog identities survive repeated staging; new identities receive distinct catalog and permission entries |
| Preservation | Unrelated native pages, application streams, catalog payloads and existing permission bytes remain retained; obsolete allocated pages require separate compaction |
| Bounds | Shared project/expansion limits; encoded native input and metadata limits; default 64 MiB recovery budget; cancellation in parsing, allocation, index traversal and writing |
| Commit | Explicit Save; same-source replacement uses Core's guarded atomic commit, including displaced-source validation and rollback on mismatch |
| Unqualified | General table/schema edits, new form/report modules, designer/event authoring, new application creation without an existing carrier, signed/protected edits, signature validation and other native generations |

Independent Access proof covers four application/designer files, ordinary class/standard module source and growth, first projects in existing native applications, and existing form/report code-behind. Access reopens output, reads exact edited source and retained host events, verifies ordinary versus host module inventories, and compiles in VBE with macros disabled. No code is executed. Native signed-project fixtures have not been qualified; unknown VBA/signature carriers are preserve-only.

## Unchanged preservation

Same-family/profile save retains the entire loaded snapshot without byte changes, including unknown pages, application payloads, linked metadata, compiled content and protection/signature carriers. Header-only and protected loads can preserve their snapshots without decoding the catalog. A loaded path is checked against its accepted SHA-256 before saving; a changed source fails before destination output. This is byte preservation, not signature validity, decryption or security-policy qualification. Signature validation remains unqualified.

Path saves stage output beside the destination and commit atomically with an explicit conflict policy. The default is `FailIfExists`. Replacing the same unchanged source path requires no write. Caller streams remain open; seekable output rewinds and truncates. Unsupported assessment and pre-cancellation preserve the previous destination. A caller-provided stream can remain partially written on an I/O failure; it has no atomic replacement mechanism.

Independent Access reopening and object re-export compare both file families and both simple/rich application corpora after a byte-identical save. Fixtures contain inert AutoExec/StopMacro, VBA, bound forms/reports and ACE embedded/data macros. Access automation forces macros disabled. No module, macro, event, link or data-macro action is invoked by the library.

## Native creation contract

New table models produce unprotected Jet 4 or ACE 12 files without reading seeds/templates, shipping blank databases, calling Microsoft Access/DAO, or adding a runtime dependency. The writer generates headers, system catalog/permission records, property maps, chained table definitions, global/table/column usage maps, packed rows, long-value chains and linked index leaf/branch pages. A complete bounded creation plan is assessed before destination I/O.

| Area | Qualified creation boundary |
| --- | --- |
| Schema | Empty and multiple-table databases; 1–255 fields per table; native rows up to 4060 bytes; up to 32 indexes including generated relationship indexes |
| Values | Boolean, Byte, Int16, Int32/AutoNumber, Currency, Single, Double, ordinary Date/Time, GUID, ShortText, LongText, Binary and Decimal precision up to 28 with declared scale; Unicode values retain exact text |
| Exactness | Decimal values must fit precision/scale without rounding; ordinary Date/Time must round-trip through native OLE dates at the same ticks; unpaired Unicode surrogates and unsupported representations are rejected |
| Omitted/null | Omitted Yes/No becomes false; explicit null Yes/No or AutoNumber is rejected. Omitted sequential AutoNumber is allocated in the output without changing the model. Other omissions become null; text/binary retain empty versus null |
| AutoNumber | One positive sequential Int32 field per table, seed ≥1, explicit values advance the persisted counter; increment/custom generator profiles are unqualified |
| Indexes | Ascending primary, unique, ordinary and composite indexes over Byte, Int16, Int32/AutoNumber and bounded ShortText; duplicate primary/unique keys and null primary members are rejected |
| Collation | Code page 1252 and General legacy sort order 1033. Qualified text keys use ASCII letters, digits, underscores and spaces, case-insensitive weights and trailing-space equivalence. Indexed Unicode outside this subset and keys over 510 bytes are rejected |
| Names | Table and relationship catalog keys require the qualified text subset. Columns/index names can be Unicode. Access-forbidden punctuation/control characters and leading/trailing whitespace are rejected; MSys table names are reserved |
| Relationships | Enforced single-field references to a unique parent index, including self references; nullable child references are allowed. Cascades and composite relationship authoring are unqualified |
| Properties | DatabaseTitle persists as AppTitle; authored text columns allow zero-length strings. AccessVersion is initialized by Access when it creates its application carrier, rather than guessed by the native table writer |
| Resources | `MaxOutputBytes` bounds planned pages (default 64 MiB). Unsupported models fail assessment; cancellation reaches plan construction and writing. The plan is invalidated by model mutations |

Fresh application/VBA creation, saved query authoring, designer/macro authoring, descending and other scalar index key codecs, modern complex/rich-text/calculated/Int64/extended-date fields, encryption and signatures remain unsupported. Existing native VBA applications use the persistence contract above; general row/schema editing and conversion remain separate work.

The portable `--create` example produces empty, multi-table and boundary databases. Independent DAO checks both file families with 255 ordinary fields, 255 long-value fields, 6000 wide rows beyond the inline allocation map, inline/single/chained long-value boundaries, scalar precision and primary/text/composite index seeks. Microsoft Access opens separate owned copies without repair, rejects duplicate keys and orphan references, updates/appends/deletes rows and re-saves. DAO and OfficeIMO independently reopen those saved files and verify the retained schema, relationships, values and edits. Fragmented/deleted/overflow input reading remains covered by the independent reader corpus; fresh creation packs new rows rather than manufacturing fragmentation.

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

`New-AccessProfileCorpus.ps1` creates the profile/protection corpus under `Fixtures/Profiles`. DAO rejects an open without the synthetic password and independently reopens the protected files with it. OfficeIMO reports protection as not assessed and does not decrypt these files. Modern feature definitions raise physical compatibility gates independently of the producer's application version. Jet 3 creation through the installed engine fails with “Could not find installable ISAM”; a genuine legacy producer remains required. Signature validation and additional protection variants remain unqualified.

`New-AccessReaderCorpus.ps1` produces [the scalar/complex corpus](../OfficeIMO.Access.Tests/Fixtures/Readers/manifest.json), including deleted/grown rows, composite indexes/relationships, Unicode/Memo/OLE, exact Currency/Decimal boundaries, rich text, lookup properties, multivalued text and attachment files. The installed ACE OLE DB engine defines Decimal precision/scale in this validation-only producer; DAO independently reopens the files read-only and exports schema, rows, complex children and decoded attachment bytes. Neither engine is a product dependency.

`New-AccessGenerationCorpus.ps1` produces [the generation corpus](../OfficeIMO.Access.Tests/Fixtures/Generations/manifest.json). Access explicitly creates format 9 and 10 MDB files; independent read-only DAO observations retain their values and saved query text. Synthetic self/password links are inventory-only. Separate DAO-produced ACE files cover calculated representations, Int64 and extended dates. The installed producer persists the extended date at scale zero; a controlled 42-byte mutation creates the seven-digit fractional probe, and an independent read-only DAO consumer confirms all fractional digits. The manifest distinguishes this controlled mutation from an Access-produced value. Native reader tests compare the persisted scale and do not silently round it.

`Test-NativeBootstrap.ps1` is a negative control: it writes a header and page-type skeleton without copying a seed file. DAO rejects both MDB and ACCDB controls (`0x800A0C0F`). This proves that header recognition is insufficient; it does **not** prove template-free native creation. No native writer is enabled by this result.

`Test-NativeCreation.ps1` runs the public `AccessDocument.Create`/`Save` examples, independent DAO/Access qualification and the OfficeIMO reader over Access-resaved files. The fixed feasibility fixtures under [Fixtures/Native](../OfficeIMO.Access.Tests/Fixtures/Native/manifest.json) remain earlier format evidence; their generator has been replaced by the public creation route. Header-derived SID masking uses the generated creation date and password region; only the observed unprotected profile is qualified.

`New-AccessDesignerCorpus.ps1` produces [the designer corpus](../OfficeIMO.Access.Tests/Fixtures/Designer/manifest.json), with independent text exports of bound forms/reports, module source, StopMacro and ACE embedded/data macros. Its fresh owned Access instance disables macro execution and imports synthetic definitions. `Test-NativePreservation.ps1` reopens and re-exports unchanged copies. Exact stream hashes and whole-file bytes are checked separately from qualified typed metadata. Direct VBA-signature validation and ACCDC distribution-package handling remain unqualified.

## Engine alternatives and deployment

The selected product boundary remains managed native codecs with no new external runtime dependency. The installed Windows Access/DAO engine is an independent qualification tool. Adopting ACE/DAO as an optional product provider would require a separate DbaClientX integration, explicit dependency approval, Windows/bitness/install policy and documented deployment limits. It would not satisfy template-free native creation.

The independently maintained [Jackcess](https://github.com/jahlborn/jackcess) Java library has an Apache-2.0 license and useful codec coverage; its database creation uses bundled empty-database resources, so adopting that creation path would not satisfy this program's seed-free requirement. [MDB Tools](https://github.com/mdbtools/mdbtools) documents Jet 3/4 page structures and provides independent reading/export; its libraries and tools have different LGPL/GPL boundaries. Neither engine nor its code/assets is referenced, vendored or shipped by OfficeIMO.Access. Reading format documentation does not approve a production dependency.

Primary references: [DAO database creation](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/dbengine-createdatabase-method-dao), [Access automation security](https://learn.microsoft.com/en-us/office/vba/api/access.application.automationsecurity), [Access object text import](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/application-load-from-text), [Access VBA and package signatures](https://support.microsoft.com/en-us/access/show-trust-by-adding-a-digital-signature-to-an-access-database), and [MDB Tools physical-format notes](https://github.com/mdbtools/mdbtools/blob/dev/HACKING.md).
