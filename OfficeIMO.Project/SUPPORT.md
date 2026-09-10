# Project format and operation support

This matrix describes the `OfficeIMO.Project` file API. Microsoft Project 2024 build **16.0.20326.20144**, running with `pl-PL` locale, produced the paired synthetic MPP/XML fixtures in `OfficeIMO.Project.Tests/Fixtures/Project2024`. Their manifest records provenance, redistribution basis, hashes, and format family.

## XML operations

| Operation | Contract | Independent evidence and boundary |
| --- | --- | --- |
| Read | Bounded XML parsing into typed objects; absent scalar values remain absent | Project 2024 empty, delivery, calendars, actuals, relationships, resources, and custom-field exports |
| Create | New XML using the application namespace and SaveVersion 14 | Authored hierarchy, durations, dependency, work assignment, cost, baseline, notes, and resource types reopened in Project 2024 |
| Field edit | Update modeled values in source XML | Name/Unicode notes, work/cost values, calendar references, and calendar exception dates; original unmodeled XML retained |
| Structural edit | Add/remove/reparent typed objects with stable UID references | Contract tests cover hierarchy and relationship integrity; opaque-reference risk blocks export unless loss permission is explicit |
| Unchanged save | Exact input bytes when loaded from bytes/stream/path and no edits were made | Every committed producer XML fixture |
| Edited save | Preserve unknown elements/attributes and update changed modeled fields | Producer fixture tests, extension-node tests, Microsoft Project semantic readback |
| Schema validation | New authored subset checked against the Project 2013 client schema | The SDK schema declares `/project/2007`; the verifier explicitly aliases it to the application's `/project` namespace without changing element rules |
| Calculation | Not implemented | Ordinary load/save never recalculates; schedule-affecting edits report stale stored values |
| Rendering and conversion | Not implemented | XML does not establish native views, formatting, printing, or MPP presentation fidelity |

The reader accepts `http://schemas.microsoft.com/project` and `http://schemas.microsoft.com/project/2007`. It retains the source namespace. Acceptance of a namespace does not qualify every producer or schema version. Project 2024 exports include extensions and ordering differences outside the older SDK schema; retained source XML is not automatically rewritten into that older profile.

## Typed fields and semantic boundaries

| Area | Typed contract | Preserved or unqualified behavior |
| --- | --- | --- |
| Identity and hierarchy | Task/resource/calendar/assignment UIDs and GUIDs; display IDs; task parent/children; WBS and outline fields | External linked-project identities remain unresolved; no implicit external file access |
| Tasks | Stored dates, duration/work/cost, progress/actual/remaining values, constraints, deadlines, notes, manual flag, and dependency links | No scheduling, leveling, constraint solving, or automatic summary arithmetic |
| Dependencies | FS, SS, FF, SF; positive/negative working or elapsed lag; percentage lag | Integer MSPDI lag precision; fractional values require explicit rounding permission. External links are retained without local binding |
| Calendars | Base-calendar references, weekday intervals, date exceptions, and legacy exception periods | Project 2024's duplicate legacy/modern exception representations stay synchronized. Recurrence details and working-week structures outside this model remain opaque; no calendar arithmetic |
| Resources | Work, material, and cost resource types; rates, units, calendar references, and stored values | Type values differ between COM and XML; the codec handles that mapping. Cost-resource amount fidelity on Project import is not qualified |
| Assignments | Task/resource identities, units, stored dates/work/cost/actual/remaining values, custom fields, baselines, timephased intervals | Microsoft Project may recalculate incomplete stored-value combinations on import; OfficeIMO does not infer scheduling inputs |
| Baselines | Slots 0–10; work/cost/BCWS/BCWP; task/assignment dates; task duration/fixed cost; task/assignment intervals | Resource baselines reject dates/intervals; assignment/resource baselines reject task-only fields. Unmodeled source fields remain preserved |
| Custom fields | Definitions, field IDs, aliases, values, lookup references, and legacy `ValueList` entries | Modern lookup/outline-code tables, formulas, masks, and enterprise semantics outside typed fields remain preserved; no expression evaluation |
| Timephased data | Compact stored intervals with type, UID, start/finish, unit, and raw value | No series expansion, aggregation, or type-specific scheduling/currency interpretation |
| Presentation and extensions | Safe source nodes and attributes retained with location-aware diagnostics | Retention is not semantic support or a native presentation conversion guarantee |

### Cost-resource import limitation

The tested Project build exports a cost-resource assignment of 300 to XML with a stored cost of 300, then reopens its own untouched XML with a cost of zero. The paired native MPP retains 300. OfficeIMO preserves the XML amount and emits `PROJECT_COST_RESOURCE_IMPORT`; it does not claim application amount fidelity or substitute guessed scheduling/timephased records. Work-resource costs in the delivery and actuals fixtures have separate successful readback evidence.

## Native format feasibility

Native experiments are opt-in tooling outside the runtime library. They do not make `ProjectDocument.Load` or `Save` support a native format.

| Experiment | Result | Qualification |
| --- | --- | --- |
| MPP14 detection and bounded compound inspection | Successful on the Project 2024 corpus | Container family markers, class identity, stream inventory, and bounded Core reader |
| Representative record extraction | Successful for task names, calendar names, and assignment identities | Fixture-qualified offsets/field IDs compared with the paired producer XML; not a general record decoder |
| Unchanged native bytes | Retained exactly | Byte copy only |
| Compound rewrite with unchanged streams | Reopened in Project 2024; stream bytes retained | Core directory/stream retention proof, not arbitrary native edit support |
| Controlled same-width task name edit | Project readback shows `Build` changed to `Craft` | One fixture and one record field; does not prove record growth, relocation, or structural changes |
| Seed-free minimal MPP creation | Rejected by Project | Required native record/header/reference layout remains unresolved; no claim of native writing |
| Legacy MPP8/9/12, MPT, MPX | Unqualified | No installed legacy application oracle or qualified producer corpus in this validation |

Project application automation disables macros and records semantic readback and re-export hashes. Successful automation with alerts suppressed does not prove that every repair warning was absent. Native writing needs a broader independently verified record contract before it can become a public API.

## Runtime and scale

The package targets `netstandard2.0`, `net8.0`, `net10.0`, and Windows `net472`. Contract tests run on Windows for .NET 8/10/.NET Framework 4.7.2 and on Ubuntu WSL for .NET 8/10. Local packed-package consumption and a Linux x64 .NET 8 NativeAOT executable exercise the fluent/XML/stream lifecycle. The AOT check is a bounded smoke test, not coverage of every field or input. macOS, browser execution, and full application integration on other producer builds remain unqualified.

The shared PowerForge benchmark runner measures process startup, load, validation, scalar edit, save, and independent streaming XML readback. Cases contain 1,000/10,000/100,000 tasks; a 10,000-level hierarchy; up to four predecessors per task; 100,000 timephased intervals; a 10,000-calendar inheritance chain; and 10,000 mirrored calendar exceptions. Calendar binding, cycle detection, and mirror matching use linear traversals with cancellation checkpoints. Explicit larger outline/element budgets are used where needed. Inputs are synthetic, so these measurements are not guarantees for arbitrary files with large opaque XML payloads.

The qualification budgets are 2 GiB process peak working set, 4 GiB allocated during load/edit/save, and 30 seconds median lifecycle time on the recorded reference machine. Cancellation is cooperative and checked during parsing and record traversal; it cannot interrupt an individual platform XML allocation. See [verification tooling](../Build/Project/README.md) for reproduction and evidence handling.
