# OfficeIMO.Project

OfficeIMO.Project creates, reads, edits, and saves Microsoft Project XML (MSPDI), reads a qualified modern MPP profile, and calculates task schedules through a typed document model. The normal and fluent APIs use the same objects. File operations and calculations run without Microsoft Project, COM, or a network connection.

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

Schedule-affecting mutations set `IsScheduleStale`; validation reports the stale stored task dates. `AreWorkCostTotalsStale` remains true after applying a schedule because that operation does not update work or cost totals. These flags describe edits in the current document; they do not certify the consistency of imported caches. Dates use `DateTimeKind.Unspecified` and represent local project wall time; no host timezone conversion occurs. Working durations use the project's minutes-per-day/week and days-per-month settings. Elapsed durations count continuous time. Costs are public currency amounts; the XML codec handles MSPDI monetary scaling. New documents use USD deterministically; set `Settings.CurrencyCode` to the intended currency. Loaded documents retain their declared currency or its absence.

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

Calculation is non-mutating and its result belongs to one document revision. Applying an error-free result rejects another document, stale revisions, and unfinished update scopes. It updates task dates, unstarted automatic-task remaining durations, and missing assignment endpoints; it leaves work/cost amounts unchanged. `Recalculate()` combines calculation and application. The scheduling profile covers calendar inheritance, dated work weeks and exceptions, split/overnight shifts, all four dependency kinds, working/elapsed/percentage lag, forward/backward scheduling, constraints, deadlines, summary dates, and float. Unsupported progress rescheduling, assignment delays/contours, external dependencies, and differing resource calendars produce diagnostics. If proposed task dates conflict with stored assignment dates, the result includes the proposed dates plus an error that blocks application: rescheduling those assignments and their curves is outside this profile. Review [the scheduling boundary](SUPPORT.md#scheduling-and-assignment-analysis) before relying on an imported project's calculation.

Calendar methods `AddWorkingMinutes`, `WorkingMinutesBetween`, and `GetWorkingIntervals` also work independently of the scheduler. Calendar searches have explicit day limits and cancellation support. `ProjectWorkEquation` exposes uniform work/duration/units and rate equations. `AnalyzeAssignments` reports stored assignment aggregates, actual/remaining inconsistencies, cached resource-total differences, and estimates where uniform rates apply. It never overwrites stored costs or actuals, applies dated rate tables, or levels resources.

## Read modern native files

```csharp
using var native = ProjectDocument.Load("delivery.mpp");
Console.WriteLine(native.NativeInfo?.ProducerVersion);
foreach (var task in native.AllTasks)
    Console.WriteLine($"{task.Uid}: {task.Name}, stored start {task.Start}");
native.Save("delivery-copy.mpp");
```

The native reader supports the tested MPP14 records produced by Microsoft Project 2024. Unchanged saves retain the entire input file exactly. Task/resource/assignment values, relationships, calendars, baseline scalar values, and local custom scalars have independent producer comparisons. Native curves, enterprise fields, formulas, lookups, and presentation records remain opaque. A native schedule calculation is a projection of decoded model values and reports that limitation. Any native mutation blocks save, even with allow-loss selected; native writing and conversion are separate unsupported operations. Read-password and write-reservation fixtures are rejected explicitly. Macro, signature, and embedded-content inventory reports conventional names without executing or verifying their content.

## Coverage and limits

See [the operation matrix](SUPPORT.md) for tested features, dialects, native-format boundaries, and platform evidence. The model includes tasks, resources, assignments, dependencies, calendars, baselines, custom fields, and compact XML timephased intervals. Native creation/editing, MPT, legacy MPP, MPX, rendering, and online service integration are unsupported.

Input limits bound bytes, XML characters/depth/elements/attributes, task outlines, entity counts, timephased intervals, and retained diagnostics. DTDs and external entities are prohibited. No linked project, schema, image, or other external resource is fetched. Set `ProjectLoadOptions` deliberately for unusually large trusted files; output has a separate `MaxOutputBytes` limit.
