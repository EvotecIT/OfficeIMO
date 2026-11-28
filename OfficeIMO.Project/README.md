# OfficeIMO.Project

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Project)](https://www.nuget.org/packages/OfficeIMO.Project)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Project?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Project)

Create, read, edit and schedule Microsoft Project files in .NET without Microsoft Project or COM. `OfficeIMO.Project` works with Project XML (MSPDI), native MPP and MPT files of generations 8, 9, 12 and 14, and MPX 4, through one typed model.

## Install

```sh
dotnet add package OfficeIMO.Project
```

## Quick start

<!-- powerforge:code source="Build/DocumentationExamples/Project/Program.cs" region="QuickStart" -->
```csharp
using OfficeIMO.Project;

using var project = ProjectDocument.Create();
project.Name = "Delivery";
project.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
project.Calendar = project.Calendars.AddStandardWorkingWeek();

var engineer = project.Resources.AddWork("Engineer");
var phase = project.Tasks.AddSummary("Delivery");
var design = phase.Children.Add("Design");
design.Duration = ProjectDuration.WorkingDays(3);
var build = phase.Children.Add("Build");
build.Duration = ProjectDuration.WorkingDays(5);
project.Dependencies.Add(design, build);
project.Assignments.Add(build, engineer, ProjectUnits.Percent(50));
project.Recalculate(new ProjectScheduleOptions { CalculateAssignments = true });
project.Validate().ThrowIfErrors();
project.Save("Delivery.xml");
```
<!-- /powerforge:code -->

Loading and saving keep stored schedule values. `Recalculate`, `ApplySchedule` and `CalculateSchedule` update or inspect task dates explicitly.

## What it does

- **Author projects** with tasks, summaries, durations, dependencies, calendars, resources and assignments, with chainable edits and forward references.
- **Read and edit** existing Project XML, MPP and MPX files, and keep what the model doesn't change.
- **Create native files** without a seed file, for the qualified MPP generations.
- **Calculate** task dates, float, work, cost and resource capacity, and inspect assignment totals.
- **Views and tables** for exchange, and conversion reports for anything a target can't carry.
- **Bounded input**: limits on bytes, XML size and depth, outlines and timephased data; DTDs and external entities are refused.

The [Project guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.project-guide.md) shows each of these with examples.

## Limits

Native curves, view and style fidelity and online services are outside the qualified profiles. Microsoft Project may recalculate work and cost when it opens an authored file. The [operation matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Project/SUPPORT.md) lists tested features, dialects and native-format boundaries.

## Learn more

- [OfficeIMO.Project on officeimo.com](https://officeimo.com/products/project/) and the [Project docs](https://officeimo.com/docs/project/).
- [Project guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.project-guide.md) and the [operation matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Project/SUPPORT.md).
- PowerShell: [PSWriteOffice](https://github.com/EvotecIT/OfficeIMO/blob/master/PowerShell/PSWriteOffice/README.MD) (`Export-OfficeProjectReport`).

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: `OfficeIMO.Core` only.
- License: MIT.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 3 | 0 | 0 | 0 | 0 | 0 |
| Read | 3 | 0 | 0 | 0 | 0 | 0 |
| Edit | 3 | 0 | 0 | 0 | 0 | 0 |
| Preserve | 3 | 0 | 0 | 0 | 0 | 0 |
| Inspect | 3 | 0 | 0 | 0 | 0 | 0 |
| Validate | 3 | 0 | 0 | 0 | 0 | 0 |
| Convert | 0 | 6 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Project` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
