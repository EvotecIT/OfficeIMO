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
| Calculation | Explicit date/float calculation and revision-bound application | See the scheduling profile below; ordinary load/save never recalculates |
| Native conversion | MPP14/MPT14 creation through the typed model with pre-write loss assessment | XML extensions and values outside the native writer profile require explicit loss permission; rendering is not implemented |

The reader accepts `http://schemas.microsoft.com/project` and `http://schemas.microsoft.com/project/2007`. It retains the source namespace. Acceptance of a namespace does not qualify every producer or schema version. Project 2024 exports include extensions and ordering differences outside the older SDK schema; retained source XML is not automatically rewritten into that older profile.

## Typed fields and semantic boundaries

| Area | Typed contract | Preserved or unqualified behavior |
| --- | --- | --- |
| Identity and hierarchy | Task/resource/calendar/assignment UIDs and GUIDs; display IDs; task parent/children; WBS and outline fields | External linked-project identities remain unresolved; no implicit external file access |
| Tasks | Stored dates, early/late dates, float, duration/work/cost, progress/actual/remaining values, constraints, deadlines, notes, manual flag, and dependency links | Explicit date calculation; no automatic work/cost/progress recalculation or leveling |
| Dependencies | FS, SS, FF, SF; positive/negative working or elapsed lag; percentage lag | Integer MSPDI lag precision; fractional values require explicit rounding permission. External links are retained without local binding |
| Calendars | Base references, weekday intervals, dated work weeks, date exceptions, legacy exception periods, working-time arithmetic | Project 2024's duplicate exception representations stay synchronized. Recurring exception rules are retained but block calendar calculation |
| Resources | Work, material, and cost resource types; rates, units, calendar references, and stored values | Type values differ between COM and XML; the codec handles that mapping. Cost-resource amount fidelity on Project import is not qualified |
| Assignments | Task/resource identities, units, stored dates/work/cost/actual/remaining values, custom fields, baselines, timephased intervals | Microsoft Project may recalculate incomplete stored-value combinations on import; OfficeIMO does not infer scheduling inputs |
| Baselines | Slots 0–10; work/cost/BCWS/BCWP; task/assignment dates; task duration/fixed cost; task/assignment intervals | Resource baselines reject dates/intervals; assignment/resource baselines reject task-only fields. Unmodeled source fields remain preserved |
| Custom fields | Definitions, field IDs, aliases, values, lookup references, and legacy `ValueList` entries | Modern lookup/outline-code tables, formulas, masks, and enterprise semantics outside typed fields remain preserved; no expression evaluation |
| Timephased data | Compact stored intervals with type, UID, start/finish, unit, and raw value | No series expansion, aggregation, or type-specific scheduling/currency interpretation |
| Presentation and extensions | Safe source nodes and attributes retained with location-aware diagnostics | Retention is not semantic support or a native presentation conversion guarantee |

### Cost-resource import limitation

The tested Project build exports a cost-resource assignment of 300 to XML with a stored cost of 300, then reopens its own untouched XML with a cost of zero. The paired native MPP retains 300. OfficeIMO preserves the XML amount and emits `PROJECT_COST_RESOURCE_IMPORT`; it does not claim application amount fidelity or substitute guessed scheduling/timephased records. Work-resource costs in the delivery and actuals fixtures have separate successful readback evidence.

## Scheduling and assignment analysis

Eight synthetic schedules cover delivery dependencies, task/resource calendar interactions, all link kinds, positive/negative/elapsed/percentage lag, all eight constraint types, deadlines, backward scheduling, fixed-work/fixed-units/fixed-duration tasks, baseline-rich input, and dated work weeks. Native and XML inputs are compared with Project 2024's exported start/finish, early/late dates, total/free float, and critical flags. This corpus qualifies those combinations, not arbitrary Microsoft Project schedules.

Calendar arithmetic uses local wall time, inherited exceptions, dated week overrides, split/overnight shifts, bounded searches, and cancellation. A task without an explicit task calendar follows its single effective resource calendar; an explicit task calendar intersects that resource calendar. Different calendars across assignments require independent assignment scheduling and are rejected. Recurring calendar exceptions remain unqualified.

Automatic tasks use explicit working or elapsed durations. Fixed-work tasks derive duration from stored work and positive assignment units when calendars agree. Manual tasks retain their explicit dates and diagnose conflicting dependencies/constraints. Summary dates roll up children; summaries with dependencies, inactive/placeholder dependency endpoints, and unsupported final bounds are rejected. Progress/status-date rescheduling, effort-driven changes to assignments, splits, leveling, non-flat assignment contours, and external-project scheduling are outside this profile. Retained XML flags for known unsupported inputs block calculation; native opaque scheduling records cannot be checked completely, so native results carry `PROJECT_NATIVE_SCHEDULE_PROJECTION`.

Applying a valid result initializes missing assignment start/finish values and updates remaining duration for unstarted automatic tasks. A newly authored dependency chain with a resource, calendar closures, and a short-Friday work week retains its calculated dates after Project 2024 opens and re-exports it. Changing task dates while retaining conflicting imported assignment dates/curves does not have that guarantee: the producer can reinstate the old schedule. Such conflicts produce `PROJECT_ASSIGNMENT_DATE_RECALCULATION_REQUIRED` and block application while leaving the proposed dates available for inspection. Work/cost amounts and actuals are never implicitly recalculated.

`AnalyzeAssignments` keeps assignment sums separate from resource caches. The producer corpus contains native resource cost caches that differ from both assignment sums and the producer's XML totals; the comparison report retains both observations and verifies the assignment sum before classifying a cache difference. Uniform estimates require explicit inputs. Dated/non-default rate tables and native rate profiles produce diagnostics instead of a rate estimate. Cost resources use entered amounts; material estimates assume fixed consumption units. No estimates are applied to the model. Applying task dates leaves potentially stale work/cost totals flagged.

## Native format feasibility

The public codec loads and writes the qualified MPP14 profile below. Twelve paired producer files in `Project2024`, `Project2024Semantics`, and `Project2024WorkWeeks` exercise native and XML observations; `Project2024Protection` adds independently protected inputs. The opt-in native authoring/edit verifier and Microsoft Project oracle exercise changed output independently of the reader.

| Public native operation | Contract | Boundary |
| --- | --- | --- |
| Read | Bounded Core compound reader, producer field maps, fixed/variable records, task hierarchy, resources, assignments, links, calendar inheritance/exceptions/work weeks, metadata | MPP14 from Project 2024 build 16.0.20326.20144; other producers/builds are unqualified |
| Baselines and local custom scalars | Task/resource/assignment baseline slots 0–10; task/resource text 1–30, number/flag 1–20, cost/date/duration 1–10; aliases | Baseline curves, enterprise values, formulas, and lookup tables stay opaque. Source absence differs from explicit native zero/default values; baseline duration estimate flags can be omitted by XML |
| Unchanged save and clone | Exact whole-file bytes | No native stream rewrite occurs |
| Field edits | Source-preserving mapped scalar updates, including variable-length Unicode names and metadata | Unsupported native mutations remain errors. Native notes editing, formulas, lookup tables, rate profiles, and curves are unqualified |
| Structural edits | Add/delete/reparent tasks; add/delete resources, assignments, calendars, and links; update mapped UID/GUID references and ordering | Opaque presentation and auxiliary references are retained without remapping; strict loss policy blocks these changes until the caller explicitly accepts that risk |
| New native files | MPP14 records and container created without a seed file or Microsoft Project | Requires a start date and named base project calendar. Resource calendars must be individually owned derived calendars. Unsupported model fields require explicit omission permission |
| Document templates | Read/write MPT14; `CreateFromTemplate` retains model and source content without associating the template as the destination | No identity or schedule reset; path-based `Global.mpt` operations are rejected. Application-wide template semantics are unqualified |
| Conversion | XML to native through the authoring profile; native to XML through the typed model | Feature-level reports identify omitted opaque content. Neither direction establishes presentation fidelity; validation errors block output even with allow-loss |
| Protected files | Read-password and write-reservation inputs rejected | No password/decryption API or claim about other protection variants |
| Inert inventory | Stream names/lengths, producer, conventional macro/signature/embedded-content presence | No macro execution, signature validation, embedded-content activation, or semantic presentation decoding |
| Timephased data | Retained in source bytes | Native curves are not exposed as typed intervals or inferred from scalar totals |

| Experiment | Result | Qualification |
| --- | --- | --- |
| MPP14 detection and bounded compound inspection | Successful on the Project 2024 corpus | Container family markers, class identity, stream inventory, and bounded Core reader |
| Representative record extraction | Qualified by the bounded public read profile above | Opaque records remain outside the typed contract |
| Unchanged native bytes | Retained exactly | Byte copy only |
| Compound rewrite with unchanged streams | Reopened in Project 2024; stream bytes retained | Core directory/stream retention proof, not arbitrary native edit support |
| Native field and structural edits | Project opens, saves again, and reopens the output with expected observed semantics | Name growth, task add/delete/reparent, resource/assignment changes, dependency links, calendar exceptions and bindings, identity changes, local custom scalars, and all eleven baseline slots |
| New MPP14 and MPT14 creation | Independently opened and saved again by Project 2024 | New hierarchy, resource assignment, dependency, dated calendar overrides, custom alias/value, and baseline; no source document payload is embedded in the writer |
| Legacy MPP8/9/12 and MPX | Unqualified | No installed legacy application oracle or qualified producer corpus in this validation |

Native scalar authoring includes mapped task dates/duration/work/cost/progress, resource types/rates/units, assignment identities/dates/work/cost/units, link kinds and lag, and calendar weekday intervals, date exceptions, and dated work weeks. The calculated `TotalSlackMinutes` cache has no qualified native record and follows the unsupported-field loss policy. Baseline writing covers work/cost in slots 0–10, task/assignment start and finish, and task duration. BCWS/BCWP, fixed-cost baseline values, and baseline curves are outside the writer profile. Local custom writing covers the task/resource text, number, flag, cost, date, and duration families listed above, plus aliases of at most 51 UTF-16 code units. Native date precision is six seconds; integer or floating-point precision changes are rejected.

Mapped edits retain unmodeled source stream content. Assignment actual start/finish and percent-work-complete caches also have no qualified native record and follow the unsupported-field policy. Structural edits warn about unknown references; schedule edits warn about retained curves and totals; edits to signed content report signature invalidation. New documents cannot carry native opaque content from an XML source. Assessments use the current model at each save, including values omitted by a prior allow-loss save. Save/clone does not turn those omissions into a lossless operation.

Project application automation disables macros and records semantic readback and re-export hashes. `-NativeRoundTrip` compares the observed model before and after an additional application save. Representative authoring and edit cases also run with `-ApplicationAlerts` enabled. The oracle checks selected semantic fields; it does not establish visual fidelity of views, printing, or every unmodeled record.

## Runtime and scale

The package targets `netstandard2.0`, `net8.0`, `net10.0`, and Windows `net472`. Contract tests run on Windows for .NET 8/10/.NET Framework 4.7.2 and on Ubuntu WSL for .NET 8/10. Local packed-package consumption and Linux x64 .NET 8 NativeAOT executables exercise the fluent/XML/stream lifecycle and native creation, repeated edits, baselines, templates, and XML conversion. The AOT checks are bounded smoke tests, not coverage of every field or input. macOS, browser execution, and full application integration on other producer builds remain unqualified.

The shared PowerForge benchmark runner measures process startup, load, validation, scalar edit, save, and independent streaming XML readback. Cases contain 1,000/10,000/100,000 tasks; a 10,000-level hierarchy; up to four predecessors per task; 100,000 timephased intervals; a 10,000-calendar inheritance chain; and 10,000 mirrored calendar exceptions. Calendar binding, cycle detection, and mirror matching use linear traversals with cancellation checkpoints. Explicit larger outline/element budgets are used where needed. Inputs are synthetic, so these measurements are not guarantees for arbitrary files with large opaque XML payloads.

The qualification budgets are 2 GiB process peak working set, 4 GiB allocated during load/edit/save, and 30 seconds median lifecycle time on the recorded reference machine. Cancellation is cooperative and checked during parsing and record traversal; it cannot interrupt an individual platform XML allocation. See [verification tooling](../Build/Project/README.md) for reproduction and evidence handling.
