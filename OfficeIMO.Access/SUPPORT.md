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
| ACE 14 | `3` / 4096 | Calculated field, encryption and compatibility-flag fixtures still required | Header recognition only |
| ACE 16 | `5` / 4096 | Large-number feature fixture still required | Header recognition only |
| ACE 17 | `6` / 4096 | Extended date/time feature fixture still required | Header recognition only |

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

The table corpus is checked in with independent expected-value exports. Application-object fixtures and exports are also project-owned synthetic evidence. Their manifests identify successful producers and the remaining missing variants. A UI-macro XML import probe failed in both families (`0x800A0750`); it does not establish an action-macro fixture or native macro carrier support. Action macros, signed/encrypted/password variants, and representative legacy/modern feature producers remain open qualification work.

`Test-NativeBootstrap.ps1` is a negative control: it writes a header and page-type skeleton without copying a seed file. DAO rejects both MDB and ACCDB controls (`0x800A0C0F`). This proves that header recognition is insufficient; it does **not** prove template-free native creation. No native writer is enabled by this result.

The required next experiment is `native-bootstrap-01`: encode a valid masked header, system-table definitions/rows, allocation maps, permissions and catalog index roots from logical definitions, then independently open the generated MDB and ACCDB. Follow with `native-table-01`: add a typed indexed user table and a relationship, export through DAO, and compare logical schema and values. Forms/reports/macros/modules need their own carrier decode/encode probes after that bootstrap. These are open acceptance criteria in the [Access roadmap](../Docs/ROADMAP.md#microsoft-access-document-library), not evidence of a working native writer.

## Engine alternatives and deployment

The selected product boundary remains managed native codecs with no new external runtime dependency. The installed Windows Access/DAO engine is an independent qualification tool. Adopting ACE/DAO as an optional product provider would require a separate DbaClientX integration, explicit dependency approval, Windows/bitness/install policy and documented deployment limits. It would not satisfy template-free native creation.

The independently maintained [Jackcess](https://github.com/jahlborn/jackcess) Java library has an Apache-2.0 license and useful codec coverage; its database creation uses bundled empty-database resources, so adopting that creation path would not satisfy this program's seed-free requirement. [MDB Tools](https://github.com/mdbtools/mdbtools) documents Jet 3/4 page structures and provides independent reading/export; its libraries and tools have different LGPL/GPL boundaries. Neither engine nor its code/assets is referenced, vendored or shipped by OfficeIMO.Access. Reading format documentation does not approve a production dependency.

Primary references: [DAO database creation](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/dbengine-createdatabase-method-dao), [Access automation security](https://learn.microsoft.com/en-us/office/vba/api/access.application.automationsecurity), [Access object text import](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/application-load-from-text), [Access VBA and package signatures](https://support.microsoft.com/en-us/access/show-trust-by-adding-a-digital-signature-to-an-access-database), and [MDB Tools physical-format notes](https://github.com/mdbtools/mdbtools/blob/dev/HACKING.md).
