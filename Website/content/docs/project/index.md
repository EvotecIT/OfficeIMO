---
title: "Create Project plans and reports in C#"
description: "Build typed Project plans, calculate schedules, and export Gantt charts, dependency networks, and timelines with complete C# examples."
order: 53
---

`OfficeIMO.Project` owns the typed project model and explicit schedule calculation. `OfficeIMO.Workflows` exports its report snapshots through the shared document and drawing engines. File operations do not require Microsoft Project or COM.

## Start with a complete example

| Example | What it demonstrates | Downloads |
|---|---|---|
| [Delivery Gantt with progress](/showcase/project-delivery-gantt/) | Working durations, dependencies, actual progress, status date, and a launch milestone | Project XML, PNG, SVG, PDF, Word |
| [Parallel dependency network](/showcase/project-dependency-network/) | Parallel work, dependency joins, critical status, and task references | Project XML, PNG, SVG, PDF |
| [Release milestone timeline](/showcase/project-milestone-timeline/) | Phase bars, zero-duration milestones, and report snapshots | Project XML, PNG, SVG, PDF, PowerPoint, Excel |

Each walkthrough includes the complete generating C# file, a checkout run command, rendered output, and downloadable files. The examples register the bundled regular and bold Carlito faces so text measurement and rendering use the same fonts.

## API entry points

Use `ProjectDocument.Create()` to build a plan or `ProjectDocument.Load(path)` to inspect and edit a supported file. Tasks, calendars, dependencies, resources, and assignments belong to the document model. Loading and saving retain stored values; calculation is explicit:

```csharp
var schedule = project.CalculateSchedule();
schedule.Report.ThrowIfErrors();
var view = project.CreateView(schedule, new ProjectViewOptions {
    Kind = ProjectViewKind.Gantt,
    Timescale = ProjectViewTimescale.Week
});
```

`CreateView` takes a snapshot. `ApplySchedule(schedule)` applies calculated task dates to the document when that is the intended edit. Choose `ProjectViewKind.Network` or `ProjectViewKind.Timeline` to present the same model differently.

`ProjectReportWorkflow.ToSvg`, `ToPdf`, and `ExportImages` create portable reports. `Images(view).WithQuality(...)` selects shared image density: 96 DPI for previews, 192 DPI for screens, or 300 DPI for printing. SVG text and geometry remain vector-based; enlarging an existing PNG does not create more detail.

`CreateWord` and `CreatePowerPoint` include chart images and editable supporting tables. `CreateExcel` retains typed dates and numeric values. Use `ProjectOfficeReportOptions` to select chart and table output.

## Scope and pagination

These are calculated report layouts, not reconstructions of saved native view definitions or styles. Native-format loading and saving follow the package's qualified profiles and assessment diagnostics. Inspect `Validate()` and `AssessSave()` before treating a conversion as lossless.

The examples fit on one page. Larger plans can produce multiple pages; consume every exported result and preserve its continuation references. Page dimensions, output counts, and raster limits are explicit so a large plan cannot silently become an unbounded image allocation.
