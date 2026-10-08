# OfficeIMO.Access design

This is the architecture and API direction for `OfficeIMO.Access`. The typed model, document lifecycle, operation assessment and bounded header inspection foundation are implemented; its current contract and qualification limits live in [Access support](../OfficeIMO.Access/SUPPORT.md). Native catalog decoding and writing remain open. Ordered delivery work and acceptance criteria live only in the [Access roadmap](ROADMAP.md#microsoft-access-document-library).

The product goal is to create, read, inspect, edit, preserve, and write Access files through one typed document model. Tables and rows are part of that model alongside saved queries, relationships, forms, reports, action macros, VBA modules, resources, and application metadata. Reading database rows alone does not satisfy this goal.

## Ownership and dependencies

| Capability | Owner | Boundary |
| --- | --- | --- |
| Access document model, object identities, native file lifecycle and preservation | `OfficeIMO.Access` | One model and one native codec owner; Jet/ACE layout details remain internal |
| Common VBA decoding, inspection, payload handling and signature primitives | `OfficeIMO.Core` and existing security owners | Extend existing owners; Access supplies its own storage/carrier adapter and qualification |
| Access-specific form/report definitions, action macros, expressions and event bindings | `OfficeIMO.Access` | Preserve uninterpreted content and expose operation-level limits |
| Database connections, provider SQL execution, transactions and live database concurrency | DbaClientX | An optional Access provider requires a separate dependency/host decision; no driver is required by the native document package |
| CSV, Excel, Word, PDF, HTML and Reader outputs | Existing OfficeIMO format owners and thin Access adapters | Keep conversion policy in the adapter and encoding/rendering in the destination owner |
| Report drawing and text layout | Existing Drawing/font/layout owners, consumed by an optional Access rendering adapter | No second renderer; static rendering has an explicit supported expression and data-binding profile |
| CLI, MCP, Workflows and Studio | Existing hosts | Map inputs and present engine results; do not implement Access parsing or database execution in hosts |
| Build, packaging, signing and publication | PowerForge/PSPublishModule | Follow the repository's release/version bindings and package qualification |

The native package depends on `OfficeIMO.Core` and qualified existing shared primitives. It introduces no external runtime dependency by default. ACE/Jet installation, Office COM automation, Java, native database utilities, and a general SQL execution engine are outside that dependency graph. Third-party engines may be isolated independent validation tools after license and artifact-boundary checks. Adding one to shipped code remains an explicit product/dependency decision.

Access file storage must be decoded through generation-specific database codecs. Existing ZIP/OPC and compound-file support is reusable only where the embedded structure actually uses it. Finding a VBA stream does not establish that the whole database is an Open XML package or compound document.

## Format and operation scope

The baseline targets `.accdb` ACE-family files and `.mdb` Jet 4 files used by Access 2000/2002/2003. The feasibility milestone pins actual header/feature profiles and producer builds; an extension or marketing version is insufficient to identify a physical format. Later baseline work qualifies Jet 3/Access 97 read/edit/write separately. Access 2 and earlier generations remain outside this baseline unless separately adopted.

Input detection validates native content before trusting a suffix. Creation uses a documented default physical profile selected in A00; callers can request another qualified profile explicitly. Incompatible filename/format/profile combinations fail before output, and a newer field type never silently upgrades the target generation.

Modern feature variants, including Large Number and Date/Time Extended, require their own profile evidence. Their presence must never be silently flattened to an earlier profile. `.accde`/`.mde` compiled applications, `.accda`/`.mda` add-ins, `.accdc` signed distribution packages and `.adp` projects are detected and diagnosed; their complete lifecycle is outside the initial baseline. Compiled VBA is preserved where possible and is never advertised as recoverable source.

Support is recorded per format profile, object/feature, and operation: detect, inspect, read, create, edit, preserve, same-profile write, profile conversion, extract, evaluate, and render. Use distinct states for supported, bounded subset, opaque preservation, explicit approximation, unsupported and blocked. Each claim links its fixture/oracle evidence and implementation owner. No single Boolean `SupportsAccess` represents this matrix.

## Public model

`AccessDocument` is the root and owns its native source, revisions and lazy readers. Typed collections expose `Tables`, `Queries`, `Relationships`, `Forms`, `Reports`, `Macros`, `VbaProject` and `Resources`. A stable document-local identity is distinct from display name and physical catalog/page identifiers. Name comparison follows the qualified Access profile without depending on the host's current culture.

| Public type family | Intended content |
| --- | --- |
| `AccessTable`, `AccessColumn`, `AccessIndex`, `AccessRelationship` | Schema, typed properties, key/index members and referential actions |
| `AccessRow`, table data reader and bounded row-edit methods | Typed values, nulls, stable row identity and explicit mutation; no mandatory full-table materialization |
| `AccessQuery` and query parameters | Original SQL, query kind, typed parameter declarations, references and supported editable structure |
| `AccessForm`, `AccessReport`, sections and controls | Layout, resource references, record/row/control sources, properties and event bindings |
| `AccessMacro` and typed actions | Standalone, embedded and table data macros; retained unknown action arguments |
| `AccessVbaProject`, module/reference descriptors | Metadata, source where available, opaque/compiled content and signature evidence |
| `AccessResource` and lazy payload handles | Attachments, images, OLE content and other inert bytes with provenance and resource bounds |
| `AccessOperationReport` and capability assessment | Object-specific locations, requested operation, severity, preservation/loss outcome and remediation |

Generation-specific representations project into this same model. Avoid parallel `MdbDocument`/`AccdbDocument` APIs or a publicly mutable bag of physical page numbers. Uninterpreted properties remain attached to their source object with provenance; they are not converted into falsely understood typed values.

A thin `AsFluent()` surface may compose these operations after the typed API is qualified. It must not introduce another state model or different validation/persistence rules. Public XML documentation and executable package examples establish the API contract before broad implementation expands it.

## Lifecycle and API direction

The [package README](../OfficeIMO.Access/README.md) and executable verification consumer define the implemented foundation API. The following example illustrates the intended native persistence and reading workflow; its save and native table-reading operations remain unsupported until their codec milestones are qualified.

```csharp
using OfficeIMO;
using OfficeIMO.Access;

using AccessDocument database = AccessDocument.Create(new AccessCreateOptions {
    Format = AccessFileFormat.Accdb
});

AccessTable contacts = database.Tables.Add("Contacts");
contacts.Columns.Add("Id", AccessDataType.AutoNumber);
contacts.Columns.Add("Name", AccessDataType.ShortText, maxLength: 120);
contacts.Indexes.AddPrimaryKey("PK_Contacts", "Id");
contacts.AppendRow(new AccessRowValues { ["Name"] = "Ada" });

AccessOperationReport assessment = database.AssessSave("contacts.accdb");
assessment.RequireNoLoss();
database.Save("contacts.accdb");

using AccessDocument source = AccessDocument.Load("source.mdb",
    new AccessLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
using var reader = source.Tables["Contacts"].OpenDataReader();
while (reader.Read()) {
    Console.WriteLine(reader.GetString(reader.GetOrdinal("Name")));
}
```

The lifecycle surface provides `Create`, `Load`, `LoadWithReport`, `AssessSave`, `Save`, `SaveWithReport`, `SaveCopy` and explicit-format stream/byte operations. Path and stream async overloads apply where they perform asynchronous I/O; cancellation also reaches long decoding, encoding and traversal loops. Convenience methods use the same engine and policy as report-returning methods.

Use the existing `DocumentAccessMode`, `DocumentPersistenceMode`, file-conflict and conversion-loss vocabulary. Default persistence is explicit. Dispose releases owned readers/resources without implicitly writing unless a caller deliberately selects a supported save-on-dispose association. Read-only rejects every mutation and write. Creation without a destination rejects save-on-dispose. A copy does not change the document's association or source profile.

Caller streams have a documented ownership policy, defaulting to leave-open. A lazy source must remain usable for the document's lifetime; copied buffering, seekability and maximum input sizes are explicit. Readers hold a document revision/lifetime lease and either prevent conflicting mutation or fail with a documented stale-revision error. Reading one table does not eagerly decode every table or payload.

`AccessRowValues` preserves a meaningful distinction between omitted fields and explicit null. Define mappings for Boolean, integer widths, Currency/fixed precision, Decimal, Single/Double, GUID, text and code pages, local date/time values, binary/long values, AutoNumber and complex values. Preserve timezone-free dates as such. Attachments and multivalued fields retain structured values, filenames/order and original bytes rather than becoming delimited strings. Lookup display text is distinct from the stored key.

Saved queries are document objects, not executable commands. Native table scans and bounded row edits require no SQL engine. Access SQL execution belongs to an optional provider path in [DbaClientX](https://github.com/EvotecIT/DbaClientX); stored SQL may invoke VBA functions or depend on host objects and cannot be treated as portable SQL automatically. A report renderer receives explicit caller-provided data and evaluates only its declared expression subset. Object renames use qualified reference parsing; text replacement across arbitrary SQL/VBA is not a safe reference update.

`BeginUpdate()` provides a document edit scope with explicit commit and rollback-on-dispose. It is not a live database transaction. Row edits, renames, relationship changes and deletes validate dependent objects before commit. Unknown references block structural edits when their safety cannot be established. Identity allocation, AutoNumber state, unique/foreign-key rules, validation rules and stale index/cache behavior are tested as observable contracts. Loading or saving never runs action queries, macro actions, data macros, VBA events or calculated expressions implicitly; writes that cannot preserve their required semantics are blocked.

## Preservation, writes and conversion

Keep original source identity and bytes or bounded backing storage, an object-to-native map, mutation revision and affected-object journal. Reuse untouched native storage only when the codec proves all cross-references remain valid. A byte-identical no-op save, opaque payload preservation, editable semantics and successful application reopen are separate evidence levels.

`AssessSave` identifies the target profile, changed objects, unsupported write paths, signature consequences and every known loss before writing. Reports distinguish a preserved but uneditable form from a fully modeled form, a dropped attachment from an explicit flattening, and stored query text from evaluated results. Strict loss policy is the default; permitted loss requires explicit caller policy and remains in the result. Blockers never become success-with-empty-output.

Path saves stage the complete file and commit atomically after validation. Protect the source from stale snapshots, active database writers, file replacement and sharing conflicts; native editing initially requires exclusive offline access. Never override Access's database locks or edit a live shared front end/back end in place. Match the repository's supported conflict policy and explain remaining platform/path guarantees. Stream saves document their weaker rollback boundary and do not imply filesystem atomicity.

Native creation means generation of a valid database header, catalog, allocation maps, table definitions, rows, indexes, relationships and required object metadata. Shipping an embedded blank database, patching a seed file, calling Access/ACE behind the native API, or exporting XML/JSON does not establish template-free native creation. Source-preserving editing, clean new-file creation and MDB/ACCDB conversion each require independent producer/reopen proof.

Forms, reports, action macros and VBA have separate native writer acceptance. An unchanged opaque blob can satisfy preservation but cannot satisfy typed editing or new object creation. Module-source changes require correct metadata/source encoding, reference handling and explicit invalidation of incompatible compiled caches and signatures; independently reopen and compile the result in an isolated Access oracle. No hidden VBA compilation/execution or external dependency resolution occurs during normal file operations.

## Security and host behavior

All embedded code, AutoExec/startup properties, OLE objects, links, connection strings and resources remain inert during native load, inspection, save and conversion. Default external resolution is denied. Linked-table metadata is inspected without opening network/file/database targets. Credentials are excluded or redacted from diagnostics, general-purpose exports and host logs; preserving a source connection string does not authorize connecting with it.

Apply checked arithmetic and configurable limits to pages, offsets, object/reference counts, tree depth/cycles, rows, strings, decompression, attachments, VBA and report output. Cancellation, corruption and limit failures must release handles and expose no partial destination. Recovery, if later adopted, is a separately named operation with incomplete-content evidence, never silent fallback in `Load`.

Password/encryption, legacy user-level security/workgroup metadata, VBA signatures and signed distribution packages are separate carrier/profile contracts. Existing OPC package-signature code cannot be assumed to understand Access. Access-specific adapters reuse qualified crypto/VBA primitives; tests use ephemeral software keys and synthetic certificates. Protection read, protection write, signature inspection, validation, removal and signing receive separate matrix rows. A blocked signed edit must remain blocked until an explicit qualified invalidation/resigning policy is selected.

Portable engine claims are qualified on Windows, Linux and macOS by operation and target framework. Microsoft Access/DAO and report exports are independent Windows validation oracles. Their use in opt-in validation does not make them production dependencies. Browser, trimming and NativeAOT support require separate proof and bounded-memory contracts; host registration does not prove that an operation runs there.

## Package structure and integration

Start with `OfficeIMO.Access`, `OfficeIMO.Access.Tests` and an executable public example/qualification lane. Separate semantic responsibilities into `Model`, `Native/Jet`, `Native/Ace`, `Read`, `Write`, `Queries`, `Forms`, `Reports`, `Macros`, `Vba`, `Preservation` and `Diagnostics` as implemented behavior requires them. Keep `AccessDocument` lifecycle partials small. Do not scaffold empty subsystems or a provider framework ahead of their milestone.

Add optional `OfficeIMO.Reader.Access` and conversion/rendering adapters only for delivered workflows. Base `OfficeIMO.Access` must remain installable without Word/Excel/PDF, DbaClientX, ACE or comparison tools. A DbaClientX provider or bridge uses a narrow data boundary and proves its package graph separately; it reuses the Access native owner if it needs native decoding. It must not force all DbaClientX providers into an Access document install.

The package README owns implemented public examples; `OfficeIMO.Access/SUPPORT.md` owns the qualified format/operation matrix and feasibility outcomes. Feed existing compatibility, conversion and operation catalogs from that source. Keep the roadmap as the only backlog and this document as architecture, not an implementation journal.

## Qualification contract

Use provenance-bound fixtures with producer/version/build, physical profile, feature inventory, license/redistribution basis, content hash, independently exported schema/data/object definitions and expected outcomes. Small synthetic fixtures isolate format failures; independent Access-produced and realistic sanitized applications establish interoperability. Include Unicode/code pages, null versus empty values, decimal/date boundaries, fragmented pages, deleted rows, indexes/relationships, complex fields, linked objects, forms/reports with code, macros, signed/encrypted inputs and malformed/truncated pages.

The oracle exports tables/schema through a separately implemented reader or Access/DAO; Access `SaveAsText` and native report PDF/image output supply object and rendering observations where suitable. Exported text is evidence, not a substitute native codec or the sole correctness oracle. Record query-export limitations and compare actual reopened application state where text export is insufficient. Automating an oracle requires a disposable test identity/environment, no user database mutation, no trust-policy changes, disabled startup/active content and isolated output.

Correctness CI covers bounded parsing, deterministic output for identical model/profile/options, cancellation, resource limits, edit rollback, strict loss handling, source preservation and package/API contracts. Opt-in evidence measures metadata inspection, first-row latency, complete scans, long/complex values, no-op saves, structural edits, native creation, rendering and repeated-operation retained memory. Compare equivalent work with validated output; keep host timing/memory budgets outside ordinary correctness gates.

Each delivered slice must pass an external-produced input path and an external-consumed output path appropriate to its claim. Source code, local packed package, public package, installed consumer and Studio/browser runtime are recorded independently. A missing oracle or unresolved native writer stays visible in the support matrix and prevents closure of the affected milestone.

## References and interpretation

- [Microsoft Access database objects](https://support.microsoft.com/en-us/access/database-basics): tables, queries, forms, reports, macros and modules share the application file.
- [Access programming](https://support.microsoft.com/en-US/Access/introduction-to-access-programming): action macros and VBA are distinct, with code bound to object events.
- [Access SQL expressions](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/sql-expressions): expression semantics include the VBA expression service; SQL text alone does not establish host-independent execution.
- [Access attachment objects](https://learn.microsoft.com/office/vba/api/access.attachment) and [multivalued fields](https://support.microsoft.com/en-us/access/create-or-delete-a-multivalued-field): these need structured modern-format mappings and explicit legacy-conversion loss.
- [Access SaveAsText](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/application-save-as-text), [LoadFromText](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/application-load-from-text) and [data/object export](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/export-all-data): useful independent observations; Microsoft documents caveats for complex query imports.
- [Access signatures](https://support.microsoft.com/en-us/access/show-trust-by-adding-a-digital-signature-to-an-access-database): direct database/VBA signing and `.accdc` package signing are distinct, with producer-version differences.
- [Access Runtime/Database Engine](https://support.microsoft.com/en-us/access/download-and-install-microsoft-365-access-runtime): provider and deployment boundaries require explicit qualification.
- [Access specifications](https://support.microsoft.com/en-gb/access/access-specifications): application limits inform profiles but do not replace lower configurable limits for untrusted files.
- Existing [VBA inspector](../OfficeIMO.Core/Internal/Vba/OfficeVbaProjectInspector.cs), [VBA canonicalization](../OfficeIMO.Core/Security/VbaSignatures/OfficeVbaProjectCanonicalizer.cs) and [document lifecycle types](../OfficeIMO.Core/Documents/DocumentLifecycle.cs) provide the initial reuse inventory.

These references establish semantics and validation routes. They do not establish OfficeIMO.Access implementation or support for an untested native layout.
