# Access support and feasibility

The current public contract is a typed model and inert header inspection. [CAPABILITIES.md](CAPABILITIES.md) and [operations.json](operations.json) are generated from `AccessCapabilities.cs`; supported native operations require independent evidence in addition to header recognition. The executable consumer in `OfficeIMO.Access.Verification` owns the example contract.

## Qualified foundation

The library builds for `netstandard2.0`, `net8.0`, `net10.0`, and Windows `net472`. The synthetic Jet 4/ACE 12 corpus comes from the installed Microsoft DAO 16 engine, independently reopens read-only, and exports typed schema and values. The manifest records byte lengths, SHA-256, producer, profile, fixture license and limitations. `OfficeIMO.Access.Tests/Fixtures` contains only synthetic project-owned data.

Foundation tests cover stream ownership and position, byte/page budgets, malformed and unknown headers, cancellation, rollback identity, detached objects, omitted versus null values, binary ownership, reader leases, typed relationships, source replacement and stale reports. They also prove that unsupported saves leave nonexistent destinations and existing streams untouched. A separate public consumer exercises these contracts without internal access.

## Profiles and qualification boundaries

| Profile | Header code / page size | Independent producer evidence | Current operations |
| --- | --- | --- | --- |
| Jet 4 MDB | `1` / 4096 | DAO 16 synthetic indexed tables, values, relationship and parameter query; Access 16 forms, report, inert VBA module and PDF output | Header inspection; no native catalog/data decoding or writing |
| ACE 12 ACCDB | `2` / 4096 | Same independent workflow | Header inspection; no native catalog/data decoding or writing |
| Jet 3 MDB | `0` / 2048 | Representative Access 97 producer fixture still required | Header recognition only |
| ACE 14 | `3`, subversion `1` / 4096 | DAO password-required database, independently reopened with a synthetic password; encryption requested at creation | Header inspection; no protection decoding or decryption |
| ACE 16 | `5` / 4096 | DAO BigInt field and independently observed `5,000,000,000` value | Header inspection; no native value decoding |
| ACE 17 | `6` / 4096 | DAO extended date/time field definition (type `26`); no precision-value fixture | Header inspection; no native value decoding |

Generation codes are physical compatibility gates, not a claim about the producing application's marketing version. Later producer applications can create older profiles. Unknown signatures/generations are rejected before a page size is assumed. Inspection diagnoses the catalog as not decoded, protection as not assessed, and structure as not validated beyond page alignment. Compiled databases, add-ins, signed packages, ADP and database templates have no qualified lifecycle in this slice.

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

`New-AccessProfileCorpus.ps1` creates the profile/protection corpus under `Fixtures/Profiles`. DAO rejects an open without the synthetic password and independently reopens the protected files with it. OfficeIMO still reports protection as not assessed and does not decrypt these files. Modern feature definitions raise the physical header compatibility gates independently of the producer's application version. Jet 3 creation through the installed engine fails with “Could not find installable ISAM”; a genuine legacy producer remains required. Signature carriers, calculated/complex fields, precision-sensitive extended date values and additional protection variants remain qualification work in A04/A09.

`Test-NativeBootstrap.ps1` is a negative control: it writes a header and page-type skeleton without copying a seed file. DAO rejects both MDB and ACCDB controls (`0x800A0C0F`). This proves that header recognition is insufficient; it does **not** prove template-free native creation. No native writer is enabled by this result.

`NativeBootstrapProbe.cs` and `Test-NativeCreation.ps1` qualify `native-bootstrap-01` and `native-table-01` for a fixed unprotected Jet 4/ACE 12 schema. The managed generator reads no seed or template. It emits the masked header, empty user slots, global/table/long-value allocation maps, system definitions/catalog/permissions rows, index leaves, two user tables, typed rows and a relationship from logical definitions. DAO independently opens both generated files, exports their persisted schema/index/relationship/value definitions, and seeks the primary index. On separate owned verification copies, it rejects duplicate primary keys and invalid foreign keys. [The native manifest](../OfficeIMO.Access.Tests/Fixtures/Native/manifest.json) records the exact bytes and observations.

This establishes feasibility, not a general production writer. The spike has a fixed schema and creation date, a restricted observed ASCII General legacy collation, an observed unprotected header field at `0x6A` whose wider semantics are unqualified, and no modern complex system catalog. Header-derived SID masking must remain consistent with the creation date and password region; the probe grants built-in permissions only inside freshly generated files. Production encoding, general schemas/values, structural edits and additional security profiles remain A05/A06/A09 work. `AccessDocument.Save` continues to fail before output.

The independent application corpus exposes different storage boundaries: Jet 4 has `MSysAccessObjects` with a storage-specific `Data` type; ACE has `MSysAccessStorage` with hierarchy/type/date metadata and an `Lv` payload. MDB Tools 1.0.0 identifies Jet's `Data` field as an unknown physical type (`0x11`), demonstrating a comparison-reader gap. Creation and text/PDF export through Access do not prove OfficeIMO carrier decoding. The Access adapters must extract and preserve form/report/macro/VBA storage before passing qualified project content to shared Core VBA/security primitives. Direct VBA signature and ACCDC distribution carriers remain separately unqualified. A04/A07 own these native carrier read/write criteria.

## Engine alternatives and deployment

The selected product boundary remains managed native codecs with no new external runtime dependency. The installed Windows Access/DAO engine is an independent qualification tool. Adopting ACE/DAO as an optional product provider would require a separate DbaClientX integration, explicit dependency approval, Windows/bitness/install policy and documented deployment limits. It would not satisfy template-free native creation.

The independently maintained [Jackcess](https://github.com/jahlborn/jackcess) Java library has an Apache-2.0 license and useful codec coverage; its database creation uses bundled empty-database resources, so adopting that creation path would not satisfy this program's seed-free requirement. [MDB Tools](https://github.com/mdbtools/mdbtools) documents Jet 3/4 page structures and provides independent reading/export; its libraries and tools have different LGPL/GPL boundaries. Neither engine nor its code/assets is referenced, vendored or shipped by OfficeIMO.Access. Reading format documentation does not approve a production dependency.

Primary references: [DAO database creation](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/dbengine-createdatabase-method-dao), [Access automation security](https://learn.microsoft.com/en-us/office/vba/api/access.application.automationsecurity), [Access object text import](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/application-load-from-text), [Access VBA and package signatures](https://support.microsoft.com/en-us/access/show-trust-by-adding-a-digital-signature-to-an-access-database), and [MDB Tools physical-format notes](https://github.com/mdbtools/mdbtools/blob/dev/HACKING.md).
