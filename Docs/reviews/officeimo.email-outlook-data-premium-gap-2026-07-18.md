# OfficeIMO.Email Outlook data premium gap audit - 2026-07-18

Branch: `codex/email-data-premium-audit`

Baseline: [PR #2105](https://github.com/EvotecIT/OfficeIMO/pull/2105) at `bc5efcb9c0f704a6fafa7a476e1ce305d66482d7`

This review treats PR #2105 as shipped product state. It assesses offline email and Outlook data workflows. SMTP, IMAP, POP3, Microsoft Graph, authentication, and delivery remain Mailozaurr responsibilities. Email HTML design and rendering remain HtmlForgeX.Email and HtmlForgeX responsibilities.

## Verdict

OfficeIMO already has a credible dependency-light Outlook artifact and mailbox engine. It is substantially beyond a replacement for MsgReader or MsgKit: it owns EML/MIME, MSG/OFT, TNEF, Mbox, ICS, VCF, PST/OST/OLM/EMLX/Maildir access, OAB v4 data, selective large-store reads, store search, semantic comparison, native export, verified PST creation and merge, and verified atomic rewrite mutation. The current [support matrix](../officeimo.email-support-matrix.md) is the source of truth for those shipped contracts.

The main premium gap is no longer another file extension. It is the programming model an Outlook COM user expects when working with data:

- a discoverable and type-safe replacement for `PropertyAccessor` and user properties;
- editable recurrence, time-zone, occurrence, and exception semantics;
- rich, composable folder queries, table projections, ordering, paging, and asynchronous traversal;
- complete typed Outlook item families, not six item kinds plus raw MAPI;
- batch property and item mutations that do not require replacing a complete `EmailDocument` manually;
- first-class categories, follow-up flags, voting, conversation, and profile/address-resolution workflows;
- recovery, split, compaction, and corruption-tolerant archive workflows where the file format permits them.

OfficeIMO should not clone the Outlook COM object model. It should provide safer value-oriented, bounded, cross-platform APIs for the data operations that made people use COM in the first place.

## Current baseline

The roadmap branch consolidates these owning surfaces into one production `OfficeIMO.Email` package while keeping
their namespaces and semantic responsibilities distinct:

| Surface | Current size | Role |
| --- | ---: | --- |
| `OfficeIMO.Email` package | 391 C# files | One assembly and NuGet for format-neutral messages/items, stores, OAB data, and mixed-artifact discovery. |
| `OfficeIMO.Email.Store` API | 187 C# files | PST/OST/OLM/EMLX/Mbox/mailbox-directory sessions, search, validation, recovery discovery, export, PST write/convert/merge, and verified PST mutation. |
| `OfficeIMO.Email.AddressBook` API | 42 C# files | Bounded OAB discovery, v4 schema/record decoding, search, validation, raw properties, contacts, and distribution-list data. |
| `OfficeIMO.Email.Data` API | 4 C# files | Thin mixed-artifact detection and dispatch to the existing individual, store, and address-book owners. |
| `OfficeIMO.Reader` and email adapters | Thin consumers | Built-in individual artifact ingestion plus optional store and address-book chunk projection. |

PR #2105 adds the current ICS/VCF, indexed Mbox, native EMLX/Maildir output, special-folder provenance, and existing-Unicode-PST mutation baseline. Its exact-head email package and packed-consumer CI lanes passed on the relevant frameworks. This audit is not a separate merge-readiness decision for that PR.

### What is already unusually strong

- Unknown and named MAPI properties survive through `MapiProperty`, including String8 source bytes and named-property identity.
- MSG/OFT coverage includes messages, appointments, contacts, tasks, journals, notes, embedded items, OLE/custom attachment storages, RTF bodies, and protected-payload handoff.
- Store reads are selective and bounded rather than materializing an entire PST or OST.
- Attachments can stay behind reopenable streams.
- Store conversion and mutation use verification, diagnostics, path identity checks, backup options, and atomic commit behavior.
- Large-store, packed-consumer, Outlook, libpff, MsgReader, MsgKit, OpenMcdf, and MimeKit oracle tests already exist in the relevant test-only lanes.
- The product has no third-party email parser runtime. `OfficeIMO.Email` directly references first-party OfficeIMO projects and Microsoft's `System.Text.Encoding.CodePages`; legacy targets also receive Microsoft compatibility packages transitively.

### Where the current public API becomes difficult

| Current shape | Practical problem |
| --- | --- |
| `IList<MapiProperty>` with numeric IDs, GUIDs, local IDs, and `object? Value` | Correct data is retained, but normal callers need protocol tables and can create a tag/type/value mismatch. There is no public typed vocabulary or property-bag operation model. |
| `EmailDocument.Properties` as `IDictionary<string, object?>` | Store, Mbox, OLM, and EMLX metadata use string keys such as `EmailStore:ItemId` and `Emlx:Flag:Flagged`. Keys, value types, mutability, and preservation behavior are undiscoverable. |
| Six known `OutlookItemKind` values backed by five nullable subtype projections | Message data uses the common model, while appointment, contact, task, journal, and note details are nullable. Invalid kind/detail combinations are possible, and distribution lists, meeting lifecycle messages, reports, posts, sharing items, and task requests have no cohesive typed model. |
| Many Outlook states are `int?` or `byte[]?` | Busy/meeting/task states and recurrence/time-zone data retain fidelity but do not provide a safe editable semantic API. |
| `EmailStoreQuery` constructor | The filter surface stops at folder, item kind, subject, sender, dates, attachments, and read state. It has no boolean expression tree, typed MAPI fields, recipient/category/importance/size filters, projection, sorting, or keyset paging. |
| Store and OAB heavy operations are synchronous | `EmailDocument` has sync/async I/O, but large store open, enumeration, search, export, conversion, merge, validation, mutation, and OAB scanning do not have a consistent asynchronous streaming contract. |
| PST mutation uses add/replace/move/delete | Safe rewrite exists, but a caller must read and replace an entire document to mark read, change categories, patch a custom property, update a flag, or edit one attachment. Copy and query-driven bulk operations are absent. |
| String folder/item IDs | IDs preserve source identity, but folder and item identifiers are interchangeable at compile time and mutation-local IDs are not represented distinctly. |

## Outlook COM workflow map

The Outlook object model exposes folders, item collections, table queries, recurrence, custom properties, categories, conversations, and address lists. Microsoft documents folder navigation and default folders through the [Folder object](https://learn.microsoft.com/en-us/office/vba/api/outlook.folder), filtering through [Items, Search, and Table APIs](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/search-and-filter/enumerating-searching-and-filtering-items-in-a-folder), arbitrary property get/set/delete through [PropertyAccessor](https://learn.microsoft.com/en-us/office/vba/api/outlook.propertyaccessor), and editable appointment/task series through [RecurrencePattern](https://learn.microsoft.com/en-us/office/vba/api/outlook.recurrencepattern).

| Common COM-era workflow | OfficeIMO status | Premium target |
| --- | --- | --- |
| Open a PST/OST, enumerate stores/folders/items | Strong offline equivalent | Add typed IDs, folder path helpers, async enumeration, and independent read cursors. |
| `GetDefaultFolder` and navigate child/parent folders | Special-folder roles exist | Add direct role/path lookup, ancestry, and folder statistics APIs. |
| `Items.Restrict`, `Find`, `Sort`, `GetTable` | Small metadata and content searches | Add a typed query AST, table projections, sort, stable paging, explain/diagnostics, and efficient table-first execution. |
| `PropertyAccessor.Get/Set/DeleteProperties` | Raw properties are retained and writable | Add typed descriptors, property bags, bulk operations, URI parsing, validation, and property availability/provenance. |
| `UserProperties` and folder field definitions | Raw named MAPI is possible | Add first-class item user properties and folder user-property definitions. Microsoft treats these as normal item/folder data in [UserProperties](https://learn.microsoft.com/en-us/office/vba/api/outlook.userproperties). |
| Read/edit recurring appointments and tasks | Opaque recurrence/time-zone blobs are retained | Decode, validate, edit, re-encode, expand occurrences, and preserve exceptions with an ICS bridge. |
| Read/write MailItem, AppointmentItem, ContactItem, TaskItem, JournalItem, NoteItem | Broad typed coverage | Add missing message-class families and enforce kind/detail invariants. |
| Read/write DistListItem and membership | OAB distribution lists are typed; MSG/PST distribution-list items remain raw contact/MAPI | Add a distinct Outlook distribution-list model and member identity types. |
| Process meeting requests/responses/cancellations and task requests | Message classes and raw data survive | Add typed lifecycle models and links to the associated appointment/task. |
| Categories, master category colors, follow-up, reminders, voting | Item category strings and some reaction payloads exist | Add typed item and store-level category catalogs, flag/follow-up/voting APIs, and validation. |
| Conversation/thread traversal | Topic/index/ID fields exist | Add a conversation graph with explainable identity and weak/strong match evidence. |
| Resolve EX/X500 recipients through address lists | OAB data and addresses are separate | Add an offline resolver that composes Store/Email/AddressBook without moving directory logic into the artifact core. |
| Copy/move/delete/patch many items | PST move/delete/replace exists in one transaction | Add copy, patch, query-driven batches, dry-run plans, conflict policy, and per-operation reports. |
| Repair, compact, split, recover deleted mail | Structural validation and indexed orphan discovery exist | Add corruption-tolerant traversal/export, soft-delete classification, split, and safe rewrite compaction. Do not call validation repair. |
| Outlook UI, inspectors, views, add-ins, account sync | Intentionally absent | Keep absent. These are not offline data-library responsibilities. |

## Competitor reality

Aspose.Email is the broadest practical commercial comparator. Its official documentation exposes PST query builders, pagination, attachment-only extraction, property updates, custom properties, bulk deletion, soft-deleted-item recovery, split, and merge in [Managing Messages in PST Files](https://docs.aspose.com/email/net/managing-messages-in-pst-files/). It also exposes editable task/calendar recurrence patterns in [Managing Recurrences](https://docs.aspose.com/email/net/managing-recurrences/), asynchronous PST create/open/merge/split in [Asynchronous Operations with PST Files](https://docs.aspose.com/email/net/asynchronous-operations-pst-files/), and distribution-list authoring in [Working with Distribution Lists](https://docs.aspose.com/email/net/working-with-distribution-lists/).

Other comparisons are narrower:

- Independentsoft PST covers common folder/item traversal, PST creation/import, existing-PST write mode, and typed message/contact/appointment/task access in its [PST .NET tutorial](https://www.independentsoft.de/pst/tutorial/index.html).
- GemBox.Email is a polished managed MSG/EML/MHTML/Mbox/ICS and transport library, but its [supported-format list](https://www.gemboxsoftware.com/email/docs/supported-formats.html) does not include PST/OST/OLM/OAB. OfficeIMO's offline-store surface is already materially broader.
- MsgReader, MsgKit, OpenMcdf, MimeKit, and libpff are valuable test oracles or narrower components, not full competitors to the combined OfficeIMO surface.

The comparison does not justify copying every commercial API. It does establish that recurrence, typed MAPI access, rich PST query/mutation, distribution lists, categories/follow-up, async archive work, split, and corruption-tolerant recovery are mainstream expectations rather than edge features.

## P0: foundations for a premium Outlook data API

### 1. Add a public typed Outlook property system

This is the first implementation slice because item semantics, query, mutation, categories, flags, user properties, and profile data all need it.

Required contract:

- `MapiPropertyKey<T>` or equivalent typed descriptors for property tags and named properties.
- Catalogs for the Microsoft `PidTag`, `PidLid`, and `PidName` properties OfficeIMO supports semantically.
- `MapiPropertyBag` over existing retained properties with `TryGet`, `Get`, `Set`, `Remove`, `Contains`, and bulk operations.
- Value/type validation before a writer sees a property.
- Preservation of unknown values and original bytes without forcing every property into the catalog.
- Property availability and origin: message, recipient row, attachment, folder, store, associated item, or address-book entry.
- Parsing and formatting of Outlook PropertyAccessor schema URIs for easier COM migrations.
- A stable policy for duplicate tags, alternate String8/Unicode representations, named-property remapping, and read-only/computed properties.
- Compatibility shims over `MapiProperties`; do not create a second property store.

Example direction:

```csharp
MapiPropertyBag mapi = document.Mapi;
string? headers = mapi.GetOrDefault(OutlookMapiProperties.TransportMessageHeaders);
mapi.Set(OutlookMapiProperties.FlagStatus, OutlookFlagStatus.Flagged);
mapi.Remove(OutlookMapiProperties.RetentionDate);

document.UserProperties.Set("CaseId", "INC-2048");
```

### 2. Decode and author Outlook recurrence and time zones

`OutlookAppointment.RecurrenceState`, `TimeZoneStructure`, and the time-zone-definition values are currently byte arrays. `OutlookTask` says whether it recurs but has no editable recurrence object. This blocks reliable calendar migration, occurrence search, task automation, and recurring-item mutation.

Microsoft specifies that `PidLidAppointmentRecur` carries the pattern, range, modified/deleted instances, and exception data in [MS-OXOCAL](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocal/5ee26cac-2c03-4b8d-8fc1-37c4bb5712dd). Exceptions can include embedded exception messages and attachments, so this must be one coherent owner rather than a convenience RRULE parser.

Required contract:

- typed daily, weekly, monthly, month-end, and yearly patterns and bounded ranges;
- task regeneration behavior;
- Windows time-zone structures and definitions with explicit unresolved-zone behavior;
- deleted and modified instances, exception attachments, and embedded exception messages;
- occurrence expansion over an explicitly bounded date range and maximum count;
- deterministic Outlook binary re-encoding and semantic reopen validation;
- loss-aware conversion to and from the existing `IcsRecurrenceRule`/`IcsTemporalValue` model;
- retention of original bytes when untouched, plus diagnostics when an edited semantic form cannot preserve vendor data.

### 3. Replace the constructor-style store filter with a typed query and table model

Do not expose Outlook DASL strings as the primary API. Build a small typed query AST that can also accept a supported DASL migration adapter later.

Required contract:

- `And`, `Or`, `Not`, equality, comparison, range, contains/prefix, existence, and bit-mask predicates;
- typed summary and MAPI fields across message, recipient, attachment metadata, folder, and store scopes;
- projection of selected columns without creating a complete `EmailDocument`;
- ascending/descending multi-key sort with a deterministic ID tie-breaker;
- stable checkpoint/keyset paging instead of only scan/result counts;
- a query plan/report identifying index/table predicates, decoded predicates, skipped/unsupported fields, scanned items, and limits;
- sync and `IAsyncEnumerable<T>` execution with cancellation;
- the same query object for search, export selection, mutation selection, Reader projection, and audit plans.

Example direction:

```csharp
EmailStoreQuery query = EmailStoreQuery.Create()
    .Where(EmailStoreFields.ReceivedAt.IsOnOrAfter(cutoff))
    .And(EmailStoreFields.Categories.Contains("Finance"))
    .And(EmailStoreFields.HasAttachments.IsTrue())
    .OrderByDescending(EmailStoreFields.ReceivedAt)
    .ThenBy(EmailStoreFields.ItemId)
    .Select(EmailStoreFields.Subject, EmailStoreFields.From, EmailStoreFields.DeclaredSize)
    .Take(500);

await foreach (EmailStoreRow row in session.QueryAsync(query, cancellationToken)) {
    // No body or attachment payload was materialized.
}
```

### 4. Complete the Outlook item family model

Add typed families where they provide stable workflow semantics, while always retaining unknown message classes and raw MAPI:

- Outlook distribution lists and members;
- meeting request, response, and cancellation items;
- task request, update, acceptance, and decline items;
- delivery/read reports and non-delivery reports;
- post and RSS items;
- sharing items;
- document items and other high-value classes found in the interoperability corpus.

Avoid one nullable property per subtype forever. Introduce a common `OutlookItemDetails` contract or discriminated result so `OutlookItemKind` and its details cannot disagree. Existing convenience properties can delegate during migration.

Replace public numeric states with enums or typed values where the Microsoft specification defines them. Keep raw properties for unknown future values rather than collapsing them to `Unknown` and losing the number.

### 5. Add composable batch mutation over the verified PST transaction

Keep the existing full semantic rewrite and verification model. Do not add a second in-place PST editor just to mimic COM.

Required operations:

- copy item and copy folder/subtree;
- patch typed fields or MAPI properties without manual full-document replacement;
- mark read/unread, set importance/sensitivity, categories, follow-up/reminder, and custom properties;
- add/remove/replace attachments and embedded items through explicit operations;
- apply a mutation to a typed query selection;
- dry-run plans with item/folder counts, estimated bytes, unsupported-property warnings, and conflicts;
- deterministic conflict policies and idempotency keys for resumable automation;
- per-operation results plus old/new stable mapping and final semantic verification;
- async commit with cancellation before the atomic replacement boundary.

One transaction should be able to express a complete archive change. Reopening and rewriting the PST once per item would be a bad COM compatibility imitation.

## P1: high-value field workflows

### 6. Make common Outlook semantics first class

- Follow-up flag state, request text, start/due/completed dates, reminders, and recipient-side follow-up.
- Master category list entries with ID, name, color, and shortcut plus item category assignment. Outlook distinguishes the item names from the store-level master category list in its [category model](https://learn.microsoft.com/en-us/office/vba/outlook/concepts/categories-and-conversations/categorize-your-outlook-items).
- Voting buttons, response value/time, and aggregate vote results.
- User properties and folder user-defined-property schemas.
- Typed reactions rather than only opaque summary/history blobs.
- Reminder state shared by messages, appointments, and tasks without duplicating three implementations.

Aspose exposes follow-up and due-date workflows as a cohesive feature in [Managing Follow-Up and Due Dates](https://docs.aspose.com/email/net/managing-follow-up-and-due-dates-for-outlook-msg-files/). OfficeIMO should own the underlying Microsoft semantics rather than copy that manager class.

### 7. Add conversation and relationship graphs

Build a bounded graph over `Message-ID`, `References`, `In-Reply-To`, normalized subject, conversation topic, conversation index, conversation ID, and meeting/task links.

The result should identify why two items were connected and distinguish authoritative from heuristic edges. It should support duplicate-aware archive views, root/children traversal, orphaned replies, cross-folder conversations, and merge/export without inventing missing messages.

### 8. Defer the protected-content boundary to a final dependency gate

The current handoff correctly detects S/MIME and retains `.p7m`/`.p7s` payloads, but a user who requires only OfficeIMO packages still needs another library to verify or decrypt the data. If the no-third-party promise includes protected-message content, the old host-handoff boundary is no longer sufficient.

This needs an explicit product-boundary decision after the Outlook-data roadmap is complete. Do not change or extract the
current PDF cryptography code in advance: separate PDF work may have changed or merged by then. At the final gate,
exercise real Outlook-authored signed and encrypted fixtures, inspect the packed transitive dependency graph, and compare
the then-current PDF implementation against the selected Bouncy Castle CMS implementation. Defer S/MIME if the real
fixtures do not produce a sufficiently small, coherent, and interoperable contract.

Gate evidence on 18 July 2026:

- the repository has S/MIME wrapper detection and raw pass-through coverage, but no genuine `.p7m`, `.p7s`, certificate,
  or Outlook-authored cryptographic fixture; the existing signed MIME samples contain placeholder signature bytes;
- a controlled `System.Security.Cryptography.Pkcs` probe created and consumed opaque signed-data, detached signatures,
  and enveloped-data on `net472`, `net8.0`, and `net10.0`, and compiled the same operations for `netstandard2.0`;
- OpenSSL 3.0 independently verified and decrypted the Microsoft-generated artifacts, while all three runnable .NET
  targets verified and decrypted OpenSSL-generated artifacts;
- Microsoft PKCS 10.0.10 is dependency-free for `net10.0`, adds two Microsoft dependencies for `net8.0`, and expands to
  a compatibility graph for `net472` and `netstandard2.0`; Microsoft PKCS 8.0.1 is dependency-free for `net472`,
  `net8.0`, and `net10.0`, while `netstandard2.0` still requires Microsoft compatibility packages;
- the former PDF-specific cryptography project was a PDF-dependent custom DER/CMS adapter and was removed when the
  neutral Security owner replaced it.

Completion evidence on 19 July 2026 closes the real-fixture gate without importing unlicensed external mail into the
repository. `ExternalOutlookSmimeCorpusTests` accepts `OFFICEIMO_EMAIL_SMIME_CORPUS`, verifies the expected artifact
hashes, and was run against [HiraokaHyperTools/smime_mail_samples commit `9e2b7f45e00c98d15ed901b9793fb3c08c20400a`](https://github.com/HiraokaHyperTools/smime_mail_samples/tree/9e2b7f45e00c98d15ed901b9793fb3c08c20400a). Its Outlook EML and binary
MSG artifacts cover signed, encrypted, and signed-then-encrypted messages with the supplied public test certificate
and private key. OfficeIMO verified, decrypted, and decrypt-then-verified all six artifacts on `net472`, `net8.0`, and
`net10.0`. A separately downloaded public Outlook 11 clear-signed mailing-list artifact was independently accepted by
OpenSSL 3.0 and exposed source-control LF normalization that OfficeIMO initially rejected; the engine now retries the
standard MIME CRLF canonical form while retaining the exact source entity, with a deterministic in-repository
regression. The external artifacts remain opt-in by path because their repositories do not grant OfficeIMO a clear
redistribution license.

The accepted implementation is one neutral `OfficeIMO.Security` package backed by Bouncy Castle. It owns bounded
CMS/DER/X.509, RFC 3161, signing, verification, and EnvelopedData operations plus vendor-neutral result/policy models.
`OfficeIMO.Pdf` and `OfficeIMO.Email` reference it directly and own only format orchestration. The former PDF-specific
cryptography package and custom CMS/DER implementation are gone, so there is no compatibility or dependency chain.

An Email-to-PDF reference, a Security package that pulls both Email and PDF, and copying CMS implementations into each
format are rejected. Public format APIs must not expose Bouncy Castle types. Target-framework support and any legacy
compatibility dependencies remain acceptance evidence, not assumptions.

The implemented Email contract includes:

- opaque and clear-signed S/MIME verification;
- enveloped-message decryption using caller-supplied certificates/keys;
- shared trust-policy callbacks and structured certificate/signature results from `OfficeIMO.Security`;
- no implicit certificate-store search by default;
- preservation of the original protected artifact and a separate decrypted `EmailDocument` result;
- format-neutral support for MIME and MSG wrappers.

Aspose treats encryption, decryption, and signature checking as data operations in its [Email Security and Encryption](https://docs.aspose.com/email/net/encrypt-decrypt-sign-email-messages/) surface. DKIM signing, transport policy, PGP, and certificate acquisition do not need to enter the first slice.

### 9. Deepen recovery and archive maintenance

OfficeIMO already discovers indexed items absent from normal folder tables and performs structural validation. The next honest steps are:

- classify normal, soft-deleted, orphaned, and structurally suspect records;
- recover/export readable items with original folder evidence and a manifest;
- tolerate isolated corrupt nodes during opt-in traversal without silently accepting them;
- split a Unicode PST by size, folder, date, or query through the existing writer;
- compact by verified rewrite while preserving semantic identity reports;
- expose capacity/health estimates before the 4 GiB per-data-tree writer limit is reached;
- produce a repair plan/report before any future repair writer exists.

Aspose separately documents [corruption-tolerant item discovery](https://docs.aspose.com/email/net/read-corrupted-pst-ost-files/) and split/recovery operations. OfficeIMO can differentiate with stricter diagnostics, manifests, and atomic verified output.

### 10. Compose offline identity resolution

`OfficeIMO.Email.AddressBook` already owns OAB entries and membership, while store items retain EX/X500 addresses and Entry IDs. Add a thin workflow API that:

- resolves a recipient against one or more OAB lists with exact and normalized evidence;
- maps EX/X500/proxy values to SMTP where the offline snapshot proves it;
- reports ambiguity, stale snapshots, and unresolved entries;
- expands distribution lists only when requested and within depth/member limits;
- never implies live Exchange resolution or freshness.

Keep the resolver in the AddressBook/identity workflow area rather than the root artifact model. It can ship in the
same dependency-free `OfficeIMO.Email` package without making ordinary `EmailDocument` behavior depend on an OAB
session.

## P2: premium breadth after the foundations

### Outlook profile, store configuration, and address-book breadth

Keep three existing boundaries distinct:

- An `OfficeIMO.Email.Profile` API area, if justified by real consumers, can own Outlook autocomplete caches (`.nk2` and supported modern Stream_Autocomplete data) with discovery, validation, search, merge, and deduplication. It should become a separate package only if a real dependency or platform boundary requires one.
- `OfficeIMO.Email.Store`, using reusable semantic codecs from `OfficeIMO.Email`, owns store-resident master category data, views, search-folder criteria, reminders, folder user-property definitions, rules, and other folder-associated information. Store already exposes associated items; do not parse them again in Profile. Microsoft stores the master category list in a Calendar-folder associated message in the [Category List specification](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxocfg/eb7cac90-6200-4ac3-8f3c-6c808c681c8b), and views are part of [folder-associated information tables](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/folder-associated-information-tables).
- `OfficeIMO.Email.AddressBook` owns OAB manifests, LZX-compressed downloads, differential patches, and older v2/v3 components.

Decision: defer a general Profile API and package. No current OfficeIMO or Mailozaurr consumer establishes enough `.nk2` or
modern `Stream_Autocomplete` demand to justify another public contract. Store and AddressBook extensions
stay with their current parsers and models; revisit Profile only with a concrete autocomplete consumer and real fixture
set.

### Additional archive formats and maintenance

- OLM authoring, plus richer OLM semantic projection and fidelity through the existing verified OLM-to-PST converter.
- MHTML/MHT read/write as a MIME archive format if real Mailozaurr/OfficeIMO consumers need it.
- PST ANSI output only if a real legacy consumer still requires it; Unicode should remain the default.
- Higher-capacity PST data-tree writing, then large-output split planning.
- OST output should remain out of scope unless a non-synchronizing offline artifact use case and interoperable writer proof emerge.

### Attachment-aware archive analysis

Content search currently handles bodies, recipients, and attachment names. A premium eDiscovery workflow should optionally extract searchable attachment text through an OfficeIMO-owned adapter:

- `OfficeIMO.Reader` and document-format packages perform extraction;
- `OfficeIMO.Email.Store` supplies attachment streams and selection;
- an orchestration package owns indexing, checkpoints, and result provenance;
- unknown/binary/encrypted attachments remain explicit, not silently empty.

Do not make Store depend on every OfficeIMO document format.

### Unified artifact discovery

The stable owners now have a thin `OfficeIMO.Email.Data` facade inside the unified package:

```csharp
EmailDataOpenResult result = EmailDataArtifact.Open(path, options);
```

It detects and returns a typed `EmailDocument`, `IcsDocument`, `VCardDocument`, `EmailStoreSession`, or
`OfflineAddressBookSession`. The facade dispatches to the existing owners and contains no alternate parsers, models,
or mutation logic. An explicit expected kind resolves ambiguous or extension-free sources.

## API consistency rules

Every new public workflow should follow these rules:

1. **Non-destructive by default.** Open/read/query never modify a source. Mutation requires an explicit transaction or output destination.
2. **One property brain.** Typed fields, raw MAPI, queries, mutation, MSG/TNEF, PST, OAB, and Reader share property descriptors and codecs.
3. **One recurrence brain.** MSG, PST, tasks, appointments, and ICS bridge through one Outlook recurrence/time-zone owner.
4. **Sync/async parity for real I/O.** Async APIs must propagate cancellation and avoid wrapping synchronous whole-store work in `Task.Run` internally.
5. **Bound every retained or scanned dimension.** Inputs, items, folders, properties, bytes, attachments, occurrences, query results, sort runs, and diagnostics need explicit limits.
6. **Report degradation.** Every conversion, projection, recovery, query fallback, and mutation returns stable diagnostics and source locations/IDs.
7. **Preserve unknown data.** Typed convenience must not discard unknown message classes, property values, extension components, or source bytes.
8. **Do not expose magic numbers as the happy path.** Raw values remain available, but normal code uses discoverable descriptors and enums.
9. **Use typed identifiers.** Folder, item, source, destination, and mutation-local IDs should not all be interchangeable strings.
10. **Plan before expensive writes.** Large exports, splits, merges, and mutations expose dry-run estimates and conflict policy.
11. **Keep package boundaries honest.** Store owns containers and store-resident associated data; Email owns message/item semantics; AddressBook owns OAB; Profile owns only true profile/cache formats; Reader owns extraction. If S/MIME is accepted at the final gate, one neutral `OfficeIMO.Security` owner uses the selected CMS dependency without exposing vendor types; Email and PDF keep only format orchestration.
12. **Validate the consumed artifact.** Reopen output with OfficeIMO and independent or native oracles where available; unit tests of internal helpers are not enough.

## Dependency promise

The public promise should be:

> Offline email and Outlook data processing without Outlook, COM automation, native mail-store libraries, or third-party email parsers.

That is stronger and more accurate than claiming a literal single-assembly dependency. Consumers can reference only the
OfficeIMO packages that own the artifacts they process; OfficeIMO may use framework, Microsoft, or one deliberately
selected implementation dependency where the capability requires it. Avoid chains of OfficeIMO adapter packages and
unrelated transitive libraries. Reader-based attachment extraction remains optional. S/MIME uses the one neutral
Security owner, so a Store consumer never needs PDF or multiple cryptography packages.

## Recommended implementation order

- [x] Publish the typed MAPI property vocabulary and property-bag API, initially covering every property OfficeIMO already reads or writes.
- [x] Route MSG, TNEF, PST, OAB, typed Outlook projections, and semantic comparison through that single vocabulary.
- [x] Add first-class user properties, categories, follow-up/reminder, and voting semantics on the property foundation.
- [x] Implement Outlook appointment/task recurrence, exceptions, time zones, bounded expansion, and ICS conversion reports.
- [x] Introduce typed store IDs, folder navigation helpers, query AST, table projections, sorting, stable paging, and async enumeration.
- [x] Add missing typed Outlook item families, starting with distribution lists and meeting/task lifecycle items.
- [x] Add batch mutation plans, copy operations, property/attachment patches, dry runs, and async verified commit.
- [x] Add conversation graphs and offline OAB-backed identity resolution.
- [x] Add corruption-tolerant recovery export, verified compaction, and query/size-based PST split.
- [x] Add Store-owned category, view, search-folder, reminder, rule, and folder user-property semantics over associated data.
- [x] Decide Profile/cache scope separately from real autocomplete consumer demand: deferred until a real autocomplete consumer exists.
- [x] Add a thin all-artifact facade only after the underlying contracts are stable: the included `OfficeIMO.Email.Data` API delegates to the three existing owners.
- [x] LAST: reassess the merged PDF signing work and freeze one neutral `OfficeIMO.Security` package backed by Bouncy Castle, with thin Email verification/decryption and PDF signing/verification consumers only where the proven contracts overlap. The bounded CMS/S/MIME contracts, tamper cases, timestamping, envelope decryption, packed dependency graphs, and opt-in exact-commit Outlook S/MIME corpus are covered without redistributing external mail.

## Current branch evidence

- `OfficeIMO.Email.Tests` passes 909 tests on `net8.0` and `net10.0`, and 896 on `net472`, including the enabled real-Outlook S/MIME corpus and Outlook-shaped MSG/TNEF clear-signed attachment verification.
- `OfficeIMO.Email.Store.Tests` passes 242 tests on `net8.0` and `net10.0`, and 231 on `net472`; four explicitly opt-in Outlook/libpff/private-corpus checks remain skipped on each target.
- `OfficeIMO.Email.AddressBook.Tests` passes 28 tests per runnable target; `OfficeIMO.Email.Data.Tests` passes three per runnable target.
- `OfficeIMO.Reader.Tests` passes 810 tests on `net8.0` and `net10.0`, and 765 on `net472`.
- `OfficeIMO.Security.Tests` passes eight CMS/X.509 tests on `net8.0` and `net10.0`; `OfficeIMO.Pdf.Tests` passes 2,987 tests on each modern target and 2,972 on `net472`.
- SDK packing succeeds for Security, PDF, unified Email, Reader.Core, every selective Reader package, and Reader.All with matching `netstandard2.0`, `net472`, `net8.0`, and `net10.0` assemblies.
- Clean isolated-cache consumers restore and execute the locally packed Security/PDF, Email/Reader.Email, and Reader.All graphs on `net472`, `net8.0`, and `net10.0`.
- The complete 127-project solution builds in Release with zero warnings and zero errors after a clean solution restore.
- `OfficeIMO.Security` owns the single Bouncy Castle dependency. PDF and Email reference Security directly; Reader.Email references only Reader.Core and Email; Reader.All references only selective Reader packages. No format-specific cryptography package, native store parser, third-party MIME parser, or third-party image library enters the graph.

## Acceptance evidence for premium claims

- Microsoft specification fixtures for every supported property, message class, recurrence pattern, exception, and time-zone structure.
- Real Outlook-authored and Outlook-reopened Windows fixtures for MSG, OFT, PST, recurrence, distribution lists, meeting/task lifecycle items, categories, user properties, flags, and custom properties.
- Semantic read-edit-write-reopen assertions; byte equality only for explicit pass-through or backup contracts.
- External corpus manifests with producer/version/provenance and expected diagnostics.
- Differential oracles against Outlook, libpff, MsgReader/MsgKit/OpenMcdf/MimeKit where each oracle is actually competent.
- Property-based and fuzz tests for MAPI values, recurrence structures, query AST serialization/planning, and malformed stores.
- Exact packed-consumer validation on `net472`, `netstandard2.0`, `net8.0`, and `net10.0`, plus Linux/macOS lanes for cross-platform surfaces.
- Large-store interruption, cancellation, resume, crash-before-commit, disk-full, path-alias, and concurrent-process tests.
- API examples written against packed OfficeIMO packages, without local binaries or competitor libraries.

## What not to build here

- SMTP, IMAP, POP3, Graph, EWS, OAuth, account synchronization, or mailbox pickup. Mailozaurr owns these workflows.
- Email template design, CSS compatibility, or HTML rendering. HtmlForgeX.Email and HtmlForgeX own those workflows.
- Outlook UI automation, add-ins, inspectors, explorers, profile creation, or identical Outlook editors.
- A line-for-line COM compatibility facade with 1-based collections, runtime string filters, and implicit save behavior.
- Exchange server directory resolution, free/busy, policy, retention enforcement, or recovery of content that is absent from an OST.
- A second MIME, MAPI, recurrence, store query, or document-extraction implementation in a convenience package.

## Best first delivery slice

Start with the typed Outlook property system, not PST split or another format.

A focused first PR should add the descriptor/key model, property-bag operations, PropertyAccessor URI parsing, catalogs for the properties OfficeIMO already projects, validation, unknown-property preservation, and MSG/TNEF/PST round-trip tests. It should expose the same bag on messages, recipients, and attachments while keeping `MapiProperties` as the underlying compatibility collection.

That slice immediately makes the existing product easier to use and creates the single foundation required by recurrence, rich queries, user properties, flags/categories, new item types, and batch mutation. Without it, each premium feature would add another private table of property IDs and another public API style.
