# OfficeIMO.Project

OfficeIMO.Project creates, reads, edits, and saves Microsoft Project XML (MSPDI), MPP/MPT generations 8, 9, 12, and 14, and MPX 4 through the [qualified profiles](SUPPORT.md). It calculates task schedules through the same typed document model used by the normal and fluent APIs. File operations and calculations run without Microsoft Project, COM, a proprietary seed file, or a network connection.

Load and save retain stored schedule values. Call `CalculateSchedule` to inspect calculated dates and float, and `ApplySchedule` or `Recalculate` to update task dates explicitly. Work, cost, actuals, and timephased values remain independently stored. Microsoft Project may calculate those values when it opens an authored file.

## Author a project

```csharp
using OfficeIMO.Project;

using var project = ProjectDocument.Create();
project.Name = "Delivery";
project.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
project.Settings.ScheduleFromStart = true;
project.Calendar = project.Calendars.AddStandardWorkingWeek();

var engineer = project.Resources.AddWork("Engineer");
engineer.StandardRate = 125;
var phase = project.Tasks.AddSummary("Delivery");
var design = phase.Children.Add("Design");
design.Duration = ProjectDuration.WorkingDays(3);
var build = phase.Children.Add("Build");
build.Duration = ProjectDuration.WorkingDays(5);
project.Dependencies.Add(design, build);
project.Assignments.Add(build, engineer, ProjectUnits.Percent(50));
project.Validate().ThrowIfErrors();
project.Save("delivery.xml");
```

## Use the fluent API

```csharp
using OfficeIMO.Project;
using OfficeIMO.Project.Fluent;

using var project = ProjectDocument.Create().AsFluent()
    .Info(info => info.Name("Delivery"))
    .StartsOn(new DateTime(2026, 10, 5, 8, 0, 0))
    .StandardWorkingWeek()
    .Resource("engineer", r => r.Name("Engineer").Work().StandardRate(125))
    .Summary("delivery", "Delivery", tasks => tasks
        .Task("design", "Design", t => t.Duration(ProjectDuration.WorkingDays(3)))
        .Task("build", "Build", t => t.Duration(ProjectDuration.WorkingDays(5))
            .After("design").Assign("engineer", ProjectUnits.Percent(50))))
    .End();

project.Tasks.GetByUid(3).Notes = "Edited through the normal API.";
project.Save("delivery.xml");
```

Builder aliases identify objects independently of their display names. `End()` validates pending alias references before adding relationships; it returns the underlying `ProjectDocument`.

## Inspect and edit

```csharp
using var project = ProjectDocument.Load("delivery.xml");
foreach (var task in project.AllTasks)
    Console.WriteLine($"{task.Uid}: {task.Name}, {task.Duration}");

var build = project.AllTasks.Single(task => task.Name == "Build");
build.Notes = "Ready for review";
var report = project.AssessSave();
report.ThrowIfErrors();
report.RequireNoLoss();
project.Save("delivery-edited.xml");
```

UIDs remain stable across edits and task moves; display IDs follow document order when a structural edit is saved. `Tasks` contains root tasks, `Children` contains immediate children, and `AllTasks` walks the whole hierarchy. Imported projects can include Microsoft's reserved UID 0 summary as a separate root metadata row. That row cannot be moved or used as an outline parent; ordinary root tasks remain parentless.

Use `MoveTo`, collection removal with an explicit `ProjectRemovalMode`, and `Clone` to maintain relationship integrity. `Clone(loadOptions: ...)` accepts explicit input limits for unusually large or deep projects. Objects from another document or removed objects cannot be attached through public setters. `BeginUpdate()` batches revision changes; finish the scope before saving. The document is mutable and is not intended for concurrent mutation.

## Streams and lifecycle

```csharp
using var input = File.OpenRead("delivery.xml");
using var project = ProjectDocument.Load(input);
project.Title = "Review copy";
using var output = new MemoryStream();
project.Save(output);
byte[] xml = output.ToArray();
```

Caller streams remain open. Seekable input is read from its beginning and its position is restored; non-seekable input is buffered within `MaxInputBytes`. Stream output replaces the contents of seekable writable streams. A failed stream write may leave partial output. File saves use the shared atomic file writer and configurable existing-file policy. Explicit save is the default; inherited `DocumentLoadOptions` control access and persistence policy. Disposal does not implicitly save unless automatic persistence was selected.

`LoadAsync` and `SaveAsync` support asynchronous stream/file I/O. Parsing and serialization are CPU work with cancellation checkpoints. `Parse` accepts an XML string and `ToXml` returns one. Use byte or stream operations when unchanged source-byte preservation matters.

## Preservation and loss

An unchanged byte-loaded document saves byte-for-byte by default. For edited documents, the codec updates modeled values in retained source XML, including unknown safe elements and attributes. `ReadDiagnostics` identifies preserved unmodeled content by location. Preservation does not mean that OfficeIMO understands or validates that content.

Structural edits can invalidate references inside unmodeled structures. `AssessSave()` reports this risk and the default loss policy blocks the save. A caller that has reviewed the findings can explicitly choose `new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }` from `OfficeIMO.Core`. Validation errors still block output. Fractional dependency lag also requires explicit loss permission because MSPDI writes integer tenths of a minute or integer percent.

Schedule-affecting mutations set `IsScheduleStale`; validation reports the stale stored task dates. Applying a date-only schedule leaves `AreWorkCostTotalsStale` set. Applying a complete assignment calculation updates the work/cost totals and clears that flag when every relevant task and assignment is covered. These flags describe edits in the current document; they do not certify the consistency of imported caches. Dates use `DateTimeKind.Unspecified` and represent local project wall time; no host timezone conversion occurs. Working durations use the project's minutes-per-day/week and days-per-month settings. Elapsed durations count continuous time. Costs are public currency amounts; the XML codec handles MSPDI monetary scaling. New documents use USD deterministically; set `Settings.CurrencyCode` to the intended currency. Loaded documents retain their declared currency or its absence.

## Calculate dates and inspect assignment totals

```csharp
var schedule = project.CalculateSchedule();
schedule.Report.ThrowIfErrors();
foreach (var task in schedule.Tasks)
    Console.WriteLine($"{task.TaskUid}: {task.Start}–{task.Finish}, float {task.TotalSlackMinutes} min");

project.ApplySchedule(schedule);
var analysis = project.AnalyzeAssignments();
foreach (var resource in analysis.Resources)
    Console.WriteLine($"{resource.ResourceUid}: assignment costs {resource.Cost}, stored resource cost {resource.StoredResourceCost}");
```

Calculation is non-mutating and its result belongs to one document revision. Applying an error-free result rejects another document, stale revisions, and unfinished update scopes. A date-only result updates task dates, unstarted automatic-task remaining durations, and missing assignment endpoints while leaving work/cost amounts unchanged. `Recalculate()` combines calculation and application. The scheduling profile covers calendar inheritance, dated work weeks and exceptions, split/overnight shifts, all four dependency kinds, working/elapsed/percentage lag, forward/backward scheduling, constraints, deadlines, summary dates, and float. Use `CalculateAssignments` for the supported work/cost, progress, delay/contour, and differing-calendar profiles. A date-only result with conflicting stored assignment dates cannot be applied. Review [the scheduling boundary](SUPPORT.md#scheduling-and-assignment-analysis) before relying on an imported project's calculation.

Calendar methods `AddWorkingMinutes`, `WorkingMinutesBetween`, and `GetWorkingIntervals` also work independently of the scheduler. Calendar searches have explicit day limits and cancellation support. `ProjectWorkEquation` exposes uniform work/duration/units and rate equations. `AnalyzeAssignments` reports stored assignment aggregates, actual/remaining inconsistencies, cached resource-total differences, and estimates where uniform rates apply. It never overwrites stored costs or actuals, applies dated rate tables, or levels resources.

## Create and edit native files

```csharp
using var native = ProjectDocument.Load("delivery.mpp");
Console.WriteLine(native.NativeInfo?.ProducerVersion);
foreach (var task in native.AllTasks)
    Console.WriteLine($"{task.Uid}: {task.Name}, stored start {task.Start}");
native.Save("delivery-copy.mpp");
```

The native codecs use separate generation profiles for Project 98 (MPP8), Project 2000–2003 (MPP9), Project 2007 (MPP12), and modern MPP14. Unchanged saves retain the entire input file exactly. Mapped field edits update source records and retain other streams. Structural changes and schedule edits carry explicit warnings about opaque references, curves, and stored totals; review `AssessSave` and select an allow-loss policy only when those limitations are acceptable. Editing an unsupported native field remains an error even with allow-loss selected.

New native documents use the same model. Set a project start and base calendar, and give each assigned work resource its own derived calendar. Save to `.mpp`, or select `ProjectFileFormat.Mpp14` for a stream:

```csharp
using var project = ProjectDocument.Create();
project.Name = "Delivery";
project.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
project.Calendar = project.Calendars.AddStandardWorkingWeek();
var engineer = project.Resources.AddWork("Engineer");
engineer.Calendar = project.Calendars.Add("Engineer", project.Calendar);
var task = project.Tasks.Add("Design");
task.Duration = ProjectDuration.WorkingDays(1);
task.Start = project.Settings.StartDate;
task.Finish = new DateTime(2026, 10, 5, 17, 0, 0);
var assignment = project.Assignments.Add(task, engineer, ProjectUnits.Percent(100));
assignment.Work = ProjectWork.Hours(8);
assignment.Start = task.Start;
assignment.Finish = task.Finish;
project.Save("delivery.mpp");
```

Native authoring covers the scalar and structural profile in [the support matrix](SUPPORT.md#native-format-feasibility). Unsupported values in a new document are reported before output; explicit loss permission can omit them. The calculated `TotalSlackMinutes` cache has no qualified native field, so saving a newly recalculated model requires accepting that omission. Dates must fit the native six-second precision, and numeric values must fit their native representation. Saving never recalculates work or cost totals.

Use `Save("delivery.mpt")` to write a document template and `ProjectDocument.CreateFromTemplate("delivery.mpt")` to create an editable, unassociated document. The template's objects, identities, and stored schedule remain intact. An explicit destination is required for the new project. `Global.mpt` is an application-wide store and is rejected by path-based operations.

File extensions select XML, MPX, project, or template output. A loaded native document retains its generation when saved to `.mpp` or `.mpt`; a new document uses generation 14. Set `ProjectSaveOptions.Format` to choose another generation or a stream format. Project formats require `.mpp`, template formats require `.mpt`, and the text formats require their matching extensions.

```csharp
var downgrade = new ProjectSaveOptions {
    Format = ProjectFileFormat.Mpp9,
    LossPolicy = OfficeConversionLossPolicy.Allow
};
var assessment = project.AssessSave("delivery-2003.mpp", downgrade);
assessment.ThrowIfErrors();
foreach (var diagnostic in assessment.Diagnostics)
    Console.WriteLine(diagnostic.Message);
project.Save("delivery-2003.mpp", downgrade);
```

Changing generations creates target records from the model and reports omitted source content. Older generations cannot represent every modern field, calendar structure, or resource type. XML-to-native conversion uses the same new-document writer. Native-to-XML conversion, including `ToXml`, reports omitted presentation records, curves/rates, custom metadata, and detected notes, calendar metadata, macros, embedded content, or signatures. It requires explicit loss permission. Review the report for the intended format with `AssessSave(new ProjectSaveOptions { Format = ProjectFileFormat.Xml })` before converting.

Native curves, enterprise fields, formulas, lookups, and presentation records remain opaque. A native schedule calculation is a projection of decoded model values and reports that limitation. Read-password and write-reservation fixtures are rejected explicitly. Macro, signature, and embedded-content inventory reports conventional names without executing or verifying their content.

## Read and write MPX

```csharp
using var exchange = ProjectDocument.Load("delivery.mpx");
exchange.AllTasks.First(task => task.Uid > 0).Notes = "First line\nSecond line";
var options = new ProjectSaveOptions {
    Format = ProjectFileFormat.Mpx4,
    MpxEncoding = ProjectMpxEncoding.Windows1252,
    MpxSeparator = ';',
    LossPolicy = OfficeConversionLossPolicy.Allow
};
exchange.AssessSave("delivery-edited.mpx", options).ThrowIfErrors();
exchange.Save("delivery-edited.mpx", options);
```

MPX 4.0/4.1 input supports English field values, numeric field tables, declared numeric/date conventions, quoted fields, multiline notes, and ANSI Windows-1252, DOS 437/850, or Macintosh Roman. New output uses MPX 4.0 and writes explicit date/time values. Unchanged output retains the original bytes; edited output uses a canonical record layout. Unsupported fields and records remain in the original bytes and require explicit loss permission before a rewrite omits them. Encodings never silently replace unsupported characters.

MPX carries currency formatting rather than a currency identity. Set `Settings.CurrencyCode` explicitly before converting an MPX input without that identity to XML. Conversion does not infer a currency from `$` or another display symbol. Calendar labels, additional baseline slots, unsupported custom fields, and features outside the MPX profile appear in the pre-write report. See [the MPX matrix](SUPPORT.md#mpx-exchange) for limits.

## Calculate work, cost, and resource capacity

Assignment calculation is explicit. It uses task and resource calendars, availability, rates, progress, and supported timephased inputs to produce a revision-bound result. Inspect diagnostics before applying it:

```csharp
project.Settings.StatusDate = new DateTime(2026, 10, 7, 8, 0, 0);
var calculated = project.CalculateSchedule(new ProjectScheduleOptions {
    CalculateAssignments = true,
    RescheduleRemainingAfterStatusDate = true
});
calculated.Report.ThrowIfErrors();
var capacity = project.AnalyzeResourceAllocation(calculated);
project.ApplySchedule(calculated);
```

Set `Settings.StatusDate` when rescheduling remaining work. `RedistributeEffortDrivenWork` and `RecalculateActualCosts` opt into their respective calculations; ordinary load/save keeps stored values. Use `CaptureBaseline` to capture a calculated baseline and `AnalyzeEarnedValue` to inspect planned value, earned value, actual cost, and variance against a selected baseline.

`CalculateLeveling` is a separate, non-mutating operation. Its options bound iterations and delay, choose priority ordering, restrict moves to available slack, and optionally split eligible work. `ApplyLeveling` applies a valid current result. See [the scheduling matrix](SUPPORT.md#scheduling-and-assignment-analysis) for unsupported combinations and tie-breaking rules.

Custom-field formulas use `CalculateCustomFields` followed by `ApplyCustomFields`. Lookup setters validate field identity and table membership. `GetOutlineCodeText` and `SetOutlineCodeValue` handle hierarchical outline values and masks. `EvaluateIndicators` applies explicit typed rules. `AddRecurringTask` creates finite occurrences from caller-provided starts; recurrence rules can generate those starts within a bounded range.

External scheduling requires an explicit `ProjectExternalProjectResolver`. OfficeIMO does not discover or fetch linked files. Resource-pool analysis uses caller-supplied document, schedule, and resource bindings; it reports shared capacity without changing other projects.

## Create views and exchange tables

```csharp
var calculated = project.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
calculated.Report.ThrowIfErrors();
var view = project.CreateView(calculated, new ProjectViewOptions {
    Kind = ProjectViewKind.Gantt,
    Timescale = ProjectViewTimescale.Week,
    BaselineNumber = 0
});
foreach (var row in view.Rows)
    Console.WriteLine($"{row.Name}: {row.Start} — {row.Finish}");
```

Views are immutable snapshots. Available layouts are Gantt, task usage, resource usage, resource histogram, network, timeline, and table. Options select tasks/resources, columns, dates, critical tasks, summaries, parent groups, baseline, and page dimensions. `Render()` creates portable drawing pages. Native Project view definitions and styles are not imported into these layouts.

Gantt, resource usage, and resource histogram default to UID and name columns, leaving more room for the time axis. Other layouts default to UID, name, start, and finish; set `Columns` to choose an explicit set. Labels, column headers, and network nodes wrap before pagination. Milestones use diamond markers, and critical status colors the chart rather than the task text. Portable pages trim unused height by default; set `FitPageHeightToContent = false` to retain the configured `PageHeight`. Content that cannot fit within the page or page budget fails explicitly.

`ExportTables` projects tasks, resources, assignments, and calendar intervals into mapped tables. The default rejects omitted project semantics; pass `allowLossyProjection: true` only after accepting the projection's diagnostic report. `ProjectDocument.ImportTables` creates a new document from explicit column mappings and import options. These tables are a data-exchange contract, not a complete project backup.

Gantt views include progress bars and summary caps. Set `ShowProgress = false` to hide progress, or set `StatusDate` to draw a reference line. Network views place dependency stages in columns and parallel tasks in separate lanes. Cards show dates, progress, and incoming/outgoing dependency references with continuation page numbers. Set `NetworkLayout = ProjectNetworkLayout.Compact` for the sequential grid. Dependency pagination can select noncontiguous rows; use each page's `RowIndices` to map its cards to source rows.

The optional [OfficeIMO.Workflows](../OfficeIMO.Workflows/README.md) package owns PDF/SVG/PNG/HTML output, charts and editable data in Word/PowerPoint/Excel reports, and CSV/Excel table transport. The base Project package depends only on the shared Core owner.

## Coverage and limits

See [the operation matrix](SUPPORT.md) for tested features, dialects, native-format boundaries, and platform evidence. Advanced calculation and report support follow the declared profiles; native curves, native view/style fidelity, and online service integration remain outside them.

Input limits bound bytes, XML characters/depth/elements/attributes, task outlines, entity counts, timephased intervals, and retained diagnostics. DTDs and external entities are prohibited. No linked project, schema, image, or other external resource is fetched. Set `ProjectLoadOptions` deliberately for unusually large trusted files; output has a separate `MaxOutputBytes` limit.
