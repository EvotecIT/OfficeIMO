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
| Format conversion | MPP/MPT 8, 9, 12, 14 and MPX 4 through the typed model with pre-write loss assessment | Extensions and values outside the target writer profile require explicit loss permission; portable report output is a separate projection |

The reader accepts `http://schemas.microsoft.com/project` and `http://schemas.microsoft.com/project/2007`. It retains the source namespace. Acceptance of a namespace does not qualify every producer or schema version. Project 2024 exports include extensions and ordering differences outside the older SDK schema; retained source XML is not automatically rewritten into that older profile.

## Typed fields and semantic boundaries

| Area | Typed contract | Preserved or unqualified behavior |
| --- | --- | --- |
| Identity and hierarchy | Task/resource/calendar/assignment UIDs and GUIDs; display IDs; task parent/children; WBS and outline fields | External identities require a caller-controlled resolver; no implicit external file access |
| Tasks | Stored dates, early/late dates, float, duration/work/cost, progress/actual/remaining values, constraints, deadlines, notes, manual flag, and dependency links | Explicit calculation and application; leveling is a separate operation |
| Dependencies | FS, SS, FF, SF; positive/negative working or elapsed lag; working/elapsed percentage lag with estimated markers | XML and MPP/MPT 8, 9, 12, 14 retain percentage time basis and estimation flags. MPX output reports loss for these flags. Integer MSPDI lag precision; fractional values require explicit rounding permission. External links are retained without local binding |
| Calendars | Base references, weekday intervals, dated work weeks, date exceptions, legacy exception periods, working-time arithmetic | Project 2024's duplicate exception representations stay synchronized. Recurring exception rules are retained but block calendar calculation |
| Resources | Work, material, and cost resource types; rates, units, calendar references, and stored values | Type values differ between COM and XML; the codec handles that mapping. Cost-resource amount fidelity on Project import is not qualified |
| Assignments | Task/resource identities, units, stored dates/work/cost/actual/remaining values, custom fields, baselines, timephased intervals | Microsoft Project may recalculate incomplete stored-value combinations on import; OfficeIMO does not infer scheduling inputs |
| Baselines | Slots 0–10; work/cost/BCWS/BCWP; task/assignment dates; task duration/fixed cost; task/assignment intervals | Resource baselines reject dates/intervals; assignment/resource baselines reject task-only fields. Unmodeled source fields remain preserved |
| Custom fields | Local definitions/scalars, bounded formulas, legacy and shared lookup tables, hierarchical outline values/masks, and explicit indicator rules | Native formula/lookup records, enterprise semantics, and native indicator presentation remain opaque |
| Timephased data | Compact stored intervals plus bounded assignment work/actual/overtime/cost and baseline interpretation | Unknown types remain stored; native curves are not decoded into typed intervals |
| Presentation and extensions | Safe source nodes and attributes retained with location-aware diagnostics | Retention is not semantic support or a native presentation conversion guarantee |

### Cost-resource import limitation

The tested Project build exports a cost-resource assignment of 300 to XML with a stored cost of 300, then reopens its own untouched XML with a cost of zero. The paired native MPP retains 300. OfficeIMO preserves the XML amount and emits `PROJECT_COST_RESOURCE_IMPORT`; it does not claim application amount fidelity or substitute guessed scheduling/timephased records. Work-resource costs in the delivery and actuals fixtures have separate successful readback evidence.

## Scheduling and assignment analysis

MPX sources with unmodeled fields or records require a fully typed model before scheduling. Opaque assignment delays, predecessor syntax, and other omitted source semantics are not silently ignored by calculation.

Mixed work/material calculations require the declared task duration to match the duration calculated from work assignments, and material intervals must fit the final task dates. Inputs requiring a new material-consumption projection are rejected before application; stored material quantities and curves are not silently rescaled.

Material actual and remaining consumption must agree with the quantity declared by assignment units and rate scale. Cost-resource calculations reconcile entered total, actual, and remaining amounts, including actual-cost curves; contradictory amounts and work quantities on cost resources block calculation. A finish-only completed milestone stays at its actual finish; a completed task with nonzero actual duration requires an actual start. Duration components and assignment actual dates must agree with their work intervals. Fixed-duration calculation derives missing actual duration from the union of work-assignment actual intervals before planning the remainder.

Working and elapsed percentage lags use their own time basis, independent of the predecessor's duration format. Estimated percentage markers remain preserved in XML and native files, but calculation requires explicit normalization to an unestimated percentage or a duration. The tested Microsoft Project build applies different arithmetic to estimated percentage formats; preserving those formats does not qualify their calculation.

Eight synthetic schedules cover delivery dependencies, task/resource calendar interactions, all link kinds, positive/negative/elapsed/percentage lag, all eight constraint types, deadlines, backward scheduling, fixed-work/fixed-units/fixed-duration tasks, baseline-rich input, and dated work weeks. Native and XML inputs are compared with Project 2024's exported start/finish, early/late dates, total/free float, and critical flags. This corpus qualifies those combinations, not arbitrary Microsoft Project schedules.

Calendar arithmetic uses local wall time, inherited exceptions, dated week overrides, split/overnight shifts, bounded searches, and cancellation. `CalculateAssignments` enables independent task/resource-calendar intersections, resource availability periods, and assignment intervals. The date-only profile still requires compatible effective calendars. Recurring calendar exceptions remain unqualified.

Automatic tasks support fixed-work, fixed-units, and fixed-duration equations through the declared assignment profile. Manual tasks retain explicit dates and diagnose conflicting dependencies/constraints. Cost-resource assignments use the final calculated task dates. Summary dates and totals roll up children; independent assignment calculation rejects direct summary-task assignments instead of omitting their costs or work. Summaries with dependencies, inactive/placeholder dependency endpoints, and unsupported final bounds are rejected. Retained XML flags for unsupported inputs block calculation; native opaque scheduling records cannot be checked completely, so native results carry `PROJECT_NATIVE_SCHEDULE_PROJECTION`.

Assigned resources require an explicit type before scheduling. Resource rates and per-use charges must be nonnegative. Work-week overrides and qualified date-range exceptions at the same calendar precedence cannot overlap inclusive dates; a child calendar can override its parent. Retained recurrence envelopes remain distinct from calculated occurrence dates. Manual assignment calculations start from the stored task start and must fit inside the stored task span.

Applying a date-only result initializes missing assignment start/finish values and updates remaining duration for unstarted automatic tasks. Conflicting imported assignment dates/curves produce `PROJECT_ASSIGNMENT_DATE_RECALCULATION_REQUIRED` and block date-only application. A result calculated with `CalculateAssignments` also applies its assignment dates, units, work/cost totals, and supported curves. Application is explicit, requires the originating document and unchanged local/external revisions, and validates the complete update before mutation. Loading and saving never trigger these calculations.

`AnalyzeAssignments` keeps assignment sums separate from resource caches. The producer corpus contains native resource cost caches that differ from both assignment sums and the producer's XML totals; the comparison report retains both observations and verifies the assignment sum before classifying a cache difference. Uniform estimates require explicit inputs. Dated/non-default rate tables and native rate profiles produce diagnostics instead of a rate estimate. Cost resources use entered amounts; material estimates assume fixed consumption units. No estimates are applied to the model. Applying task dates leaves potentially stale work/cost totals flagged.

### Advanced calculation profile

Microsoft Project build **16.0.20326.20132** produced the synthetic XML fixtures in `Project2024Advanced`. They cover differing calendars, effort-driven changes, progress, rates/material usage, contours, splits, leveling, custom fields, outline codes, recurrence, and linked projects. The fixture manifest identifies the producer and each file. Application reopens additionally exercise authored/edited assignment schedules, lookup/outline values, and applied interruptions. These observations qualify the named combinations; native record decoding remains subject to the separate native matrix.

| Area | Supported calculation | Boundary |
| --- | --- | --- |
| Assignment equations | Fixed units/work/duration, independent effective calendars, dated availability, and explicit effort redistribution | `RedistributeEffortDrivenWork` is opt-in; incomplete or inconsistent inputs produce diagnostics |
| Progress | Actual/remaining work and duration, work completion, duration completion, physical completion, overtime, stop/resume, and optional status-date rescheduling | Recorded actuals stay anchored. Positive completion percentages require actual work or duration inputs. Timephased actual overtime requires explicit actual-work intervals. Conflicting work components and unsupported or overlapping curves block calculation rather than being discarded |
| Work curves | Flat, front/back loaded, double/early/late peak, bell, and explicit custom intervals; zero-work gaps preserve splits | Turtle requires explicit producer intervals to retain its rounding. Native curves remain opaque |
| Cost | Work/material/cost resources, fixed and variable material use, dated rate-table selection, overtime rates, per-use costs, and start/end/prorated accrual | Recalculated actual costs are opt-in; unknown amounts stay unknown. Cost-resource application-import limitation still applies |
| Baselines and earned value | Capture slots 0–10, baseline work/cost curves, status-clipped PV/EV/AC, variance, CPI/SPI, and CPI-based forecast | Missing curves or inconsistent totals produce diagnostics and nullable results, not invented uniform budgets |
| Capacity and leveling | Interval-based overload analysis; lower-priority tasks move first, with descending UID breaking ties; bounded delay, optional slack bounds and remaining-work splits | Priority 1000 prevents movement. Actual work is not moved. Unresolvable constraints and exhausted limits return errors; ordinary scheduling never levels |
| External schedules | Explicit resolver, project/task identity mapping, bounded recursive resolution, dependency dates, source revision checks | No automatic file access or source mutation; unsupported external combinations are diagnosed |
| Resource pools | Explicit local-to-pool resource bindings across current calculated documents; shared dated-capacity analysis | No implicit native pool-link resolution, cross-project leveling, or synchronization of calendars/rates |

### Local custom fields and recurrence

The bounded formula evaluator supports local field references, arithmetic/comparison/logical operations, text concatenation, explicit culture, and supported summary rollups. Functions are `IIf`, `Abs`, `Round`, `Int`, `Fix`, `CStr`, `CBool`, `UCase`, `LCase`, `Trim`, `Len`, `Left`, `Right`, `Mid`, `Year`, `Month`, `Day`, `Hour`, `Minute`, `Second`, and `DateAdd`. Unsupported functions, cycles, ambiguous identities, and exhausted expression budgets produce diagnostics. Evaluation never executes arbitrary code. This is not an implementation of every Microsoft Project/VBA expression.

Lookup APIs support legacy inline value lists and modern shared flat scalar tables. Outline codes support the ten local task/resource slots, hierarchical text values, separators, numeric/upper/lower/any-character masks, and leaf/complete-path restrictions. Setters and save validation check table/value GUIDs, IDs, ownership, membership, masks, and cycles. Enterprise fields and native lookup/formula persistence remain outside the profile. Graphical indicators are explicit typed rules and portable results; native indicator tables are not reconstructed.

Recurrence expands finite daily, weekly, monthly, and yearly Gregorian patterns or explicit occurrence starts into summary/child tasks. Expansion preserves local wall time, skips nonexistent month days, and respects occurrence/search limits. Working calendars are applied during subsequent scheduling. Producer XML retains expanded occurrences and recurring markers but does not contain an editable native recurrence rule. Authored XML imports as ordinary tasks; native recurrence-rule editing is unqualified.

## Portable reports and mapped data exchange

| Operation | Contract | Evidence and boundary |
| --- | --- | --- |
| View creation | Immutable Gantt, task/resource usage, resource histogram, network, timeline, and table snapshots from a current calculation | Task/resource selection, name/critical filters, summary inclusion, parent grouping, columns, baseline dates, day/week/month buckets, and bounded page dimensions. Critical-only summary work, cost, and usage exclude noncritical descendants; hidden critical children still contribute |
| Diagrams | Core drawing pages with critical/summary/baseline styling, legends, and pagination | Gantt draws dependencies whose endpoints are visible on the same page. Network labels retain relationship type/lag and cross-page references. Native Project view/style tables are not interpreted |
| PDF/SVG/PNG/HTML | Thin `ProjectReportWorkflow` adapters over Core/Pdf/Html | Small and 75-task outputs inspected for pages, Unicode, clipping, links, and empty states. HTML contains accessible tables and horizontally scrollable diagram pages; print CSS uses configured page dimensions |
| Editable reports | Native Word/PowerPoint tables and Excel worksheets, including usage, status/groups, baseline dates, and dependencies | Reopened and exported by Microsoft Office 16.0. Large cases retain final rows through Word and PowerPoint pagination. These are editable report data, not native Project layouts |
| Table projection | Explicit mapped task/resource/assignment/calendar tables with stable identities and invariant units | Default export rejects omitted semantics; explicit lossy projection returns diagnostics. Import creates a new document and validates mappings, references, duplicates, units, and limits |
| CSV/Excel transport | Canonical CSV/Excel owners preserve quoted/Unicode/literal text; Excel values remain typed | Numeric import rejects precision loss, date kinds are explicit, and formulas are not evaluated as imported project data. No complete native-project reconstruction claim |

Report dates are local wall time. Explicit `Finish` is exclusive; inferred ranges include terminal milestones. Visible usage buckets are clipped to the selected dates, while row work totals describe the whole selected task/resource. Resource-filtered task totals include selected assignments and exclude task fixed costs; task dates and progress remain task-wide. Resource histograms display work hours, not percentage utilization. Long native PowerPoint rows are measured and paginated; an unfit row fails with a diagnostic exception rather than silently dropping text. Output page/row/cell/interval limits and cancellation bound report work.

`OfficeIMO.Project` owns the semantic views and depends only on Core. Optional PDF/HTML/CSV/Excel/Word/PowerPoint composition belongs to `OfficeIMO.Workflows`. Reports and data tables are projections, not reversible project formats. Font coverage and native Office pagination depend on the chosen fonts and application; supplied rendering profiles provide explicit font input.

## Native format feasibility

The public codecs load and write separate generation profiles. Twelve paired producer files in `Project2024`, `Project2024Semantics`, and `Project2024WorkWeeks` exercise modern native and XML observations; `Project2024Protection` adds independently protected inputs. Seven paired `Project2024Mpp12` files exercise Project 2024's Project 2007 export. Opt-in historical-fixture checks and independent readback cover the older profiles without adding a runtime dependency.

| Generation | Read and unchanged save | New files, edits, templates, and conversion | Application boundary |
| --- | --- | --- | --- |
| MPP14 / MPT14 | Project 2024 fixtures and exact original bytes | Field/structural edit, authoring, template, and conversion output reopened by Project 2024 | Other producer builds and application-wide templates remain unqualified |
| MPP12 / MPT12 | Project 2024 exports and a historical Project 2007 relationship fixture | Independent reader accepts authored and converted files; Project 2024 opens, saves, and reopens authoring, edits, and upgrade/downgrade output | Qualification applies to the tested records and producers |
| MPP9 / MPT9 | Historical Project 2000/2003 fixtures, including calendars, assignments, links, and redundant trailing metadata | Eight authored and eight producer-based lifecycle cases reopened by Project 2024; independent field comparisons cover the historical fixtures | The modern application can recalculate stored totals and remap assignment UIDs on import |
| MPP8 / MPT8 | Eight historical Project 98/2003 fixtures, including dates, durations, assignments, links, and priority | Twenty-four authoring/edit/template cases accepted by the independent reader; upgrade output opened by Project 2024 | No Project 98 application readback. The installed modern application cannot directly open MPP8 |

| Public native operation | Contract | Boundary |
| --- | --- | --- |
| Read | Bounded Core compound reader, generation-specific field maps and record storage, hierarchy, resources, assignments, links, calendars, metadata | Generation qualification is listed above. Project 98 uses bounded deferred block chains; unsupported external variable-storage layouts are rejected |
| Baselines and local custom scalars | Mapped task/resource/assignment baselines and local text, number, flag, cost, date, duration, start, and finish fields | Available slots differ by generation. Modern profiles cover baseline slots 0–10, task/resource text 1–30, number/flag 1–20, cost/date/duration 1–10, and task start/finish 1–5. Curves, enterprise values, formulas, and lookups remain opaque |
| Unchanged save and clone | Exact whole-file bytes | No native stream rewrite occurs |
| Field edits | Source-preserving mapped scalar updates, including variable-length Unicode names and metadata | Unsupported native mutations remain errors. Native notes editing, formulas, lookup tables, rate profiles, and curves are unqualified |
| Structural edits | Add/delete/reparent tasks; add/delete resources, assignments, calendars, and links; update mapped UID/GUID references and ordering | Opaque presentation and auxiliary references are retained without remapping; strict loss policy blocks these changes until the caller explicitly accepts that risk |
| New native files | MPP8/9/12/14 records and containers created without a seed file or Microsoft Project | Requires a start date and named base project calendar. Resource calendars must be individually owned derived calendars. Unsupported model fields require explicit omission permission |
| Document templates | Read/write MPT8/9/12/14; `CreateFromTemplate` retains model, generation, and source content without associating the template as the destination | No identity or schedule reset; path-based `Global.mpt` operations are rejected. Application-wide template semantics are unqualified |
| Conversion | All 100 pairs among XML, MPX, and the eight native project/template formats exercise assessment, save, and reopen | Ninety native/MPX target files also have independent reader comparisons. Representative routes into MPP12/14 pass application save/reopen. This does not establish presentation fidelity |
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
| Historical native field comparison | Fourteen historical MPP8/9/12 fixtures have no unexplained mapped-field differences | Reports retain source defaults, derived critical flags, empty summary flags, and a synthesized orphan assignment separately |

Native scalar authoring includes mapped task dates/duration/work/cost/progress, resource types/rates/units, assignment identities/dates/work/cost/units, link kinds and lag, and calendar weekday intervals, date exceptions, and dated work weeks. The calculated `TotalSlackMinutes` cache has no qualified native record and follows the unsupported-field loss policy. Baseline writing covers work/cost in slots 0–10, task/assignment start and finish, and task duration. BCWS/BCWP, fixed-cost baseline values, and baseline curves are outside the writer profile. Local custom writing covers the task/resource text, number, flag, cost, date, and duration families listed above, plus aliases of at most 51 UTF-16 code units. Native date precision is six seconds; integer or floating-point precision changes are rejected.

Mapped edits retain unmodeled source stream content. Assignment actual start/finish and percent-work-complete caches also have no qualified native record and follow the unsupported-field policy. Structural edits warn about unknown references; schedule edits warn about retained curves and totals; edits to signed content report signature invalidation. New documents cannot carry native opaque content from an XML source. Assessments use the current model at each save, including values omitted by a prior allow-loss save. Save/clone does not turn those omissions into a lossless operation.

MPP8 supports work resources, up to three working intervals per day, and its native priority bands. Priority quantization and omitted aliases are reported. MPP9 supports work/material resources; cost resources require MPP12 or later. Legacy calendar work weeks are flattened into bounded date exceptions with explicit structure/label loss. Full-day intervals use equal start/end clocks and retain 24 hours through every codec. Project 98's qualified finish profile requires stored finish and early finish to agree. Missing built-in Standard-calendar data uses the declared historical Monday–Friday default; an unknown global calendar is rejected.

The native field comparison distinguishes stored values from values an independent reader or Microsoft Project derives. For example, zero or absent slack can yield a critical flag even when the stored flag is false, and omitted progress fields can export as zero. Application import can also recalculate an incomplete progress/work/cost combination. Raw-field round trips and application-calculated values are separate observations.

Project application automation disables macros and records semantic readback and re-export hashes. `-NativeRoundTrip` compares the observed model before and after an additional application save. Representative authoring and edit cases also run with `-ApplicationAlerts` enabled. The oracle checks selected semantic fields; it does not establish visual fidelity of views, printing, or every unmodeled record.

## MPX exchange

| Area | Qualified contract | Boundary |
| --- | --- | --- |
| Dialect | MPX 4.0/4.1 input, MPX 4.0 output, numeric field tables, English field names and enumerated values | Other versions and localized field-value vocabularies are rejected or unqualified |
| Encoding | Windows-1252 ANSI, DOS 437/850, Macintosh Roman | Four independent encoding readbacks; unrepresentable output characters fail before writing |
| Framing and values | Declared delimiter, doubled quotes, numeric separators, MDY/DMY/YMD dates, English date names, 12/24-hour clocks, working/elapsed durations, rate units, currency formatting | New output uses explicit minute-resolution dates. Rewrites reject `+`, `-`, `.`, `%`, and `?` delimiters because they conflict with dependency syntax; ambiguous input links remain opaque. Source currency symbols do not establish a currency code |
| Lifecycle | Exact unchanged bytes; canonical edited output; task/resource/assignment creation, deletion, hierarchy edits, notes, calendars, and links | Seven authored and seven independently produced lifecycle cases have independent reader comparisons |
| Mapped fields | Core task/resource dates, work/cost/progress, constraints, priority, links, default baseline, and supported scalar custom fields | Task Text1–10, resource Text1–5, Number1–5, Flag1–10, Cost1–3, Duration1–3, and task Start/Finish1–5; other fields require explicit omission permission |
| Calendars | Base/resource calendars, weekly intervals, date exceptions, and bounded work-week projection | At most 250 base calendars, 250 exceptions per calendar, and three intervals per day. Projection can change calendar identity and omit labels/structure |
| Limits | 9,999 tasks/resources, 100 assignments per task, configurable input/output and entity budgets | Fractional progress rounds to the integer model with a loss diagnostic; source bytes retain the original value |
| Unknown content | Original bytes retain unmapped fields, recurrence metadata, DDE/OLE records, and other unknown records | Rewriting requires explicit loss permission; retained external records never execute or resolve files |
| Independent corpus | Eighteen English historical MPX fixtures and four encoding fixtures | Comparison retains derived remaining duration, WBS, and integer-progress differences as observations. No legacy Microsoft Project application readback |

## Runtime and scale

The package targets `netstandard2.0`, `net8.0`, `net10.0`, and Windows `net472`. Contract tests run on Windows for .NET 8/10/.NET Framework 4.7.2 and on Ubuntu WSL for .NET 8/10. Local packed-package consumption and Linux x64 .NET 8 NativeAOT executables exercise the fluent/XML/stream lifecycle and native creation, repeated edits, baselines, templates, and XML conversion. The AOT checks are bounded smoke tests, not coverage of every field or input. macOS, browser execution, and full application integration on other producer builds remain unqualified.

The shared PowerForge benchmark runner measures process startup, load, validation, scalar edit, save, and independent streaming XML readback. Cases contain 1,000/10,000/100,000 tasks; a 10,000-level hierarchy; up to four predecessors per task; 100,000 timephased intervals; a 10,000-calendar inheritance chain; and 10,000 mirrored calendar exceptions. Calendar binding, cycle detection, and mirror matching use linear traversals with cancellation checkpoints. Explicit larger outline/element budgets are used where needed. Inputs are synthetic, so these measurements are not guarantees for arbitrary files with large opaque XML payloads.

The qualification budgets are 2 GiB process peak working set, 4 GiB allocated during load/edit/save, and 30 seconds median lifecycle time on the recorded reference machine. Cancellation is cooperative and checked during parsing and record traversal; it cannot interrupt an individual platform XML allocation. See [verification tooling](../Build/Project/README.md) for reproduction and evidence handling.
