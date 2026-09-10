# Project interoperability verification

These opt-in tools produce synthetic fixtures and compare OfficeIMO output with Microsoft Project. They are outside the runtime package; the console verifier is outside the normal solution. Run from the repository root and choose new output directories so evidence is not overwritten.

## Build and run

```powershell
dotnet test OfficeIMO.Project.Tests -c Release
dotnet build OfficeIMO.Project.Verification -c Release

# Windows PowerShell 5.1 and an installed, licensed Microsoft Project are required.
powershell.exe -NoProfile -File Build/Project/New-ProjectFixtures.ps1 -OutputPath ./artifacts/project/fixtures

dotnet run --project OfficeIMO.Project.Verification -c Release -- `
    OfficeIMO.Project.Tests/Fixtures/Project2024/delivery.xml ./artifacts/project/xml-proof

powershell.exe -NoProfile -File Build/Project/Test-ProjectInteroperability.ps1 `
    -InputPath ./artifacts/project/xml-proof/authored.xml -OutputPath ./artifacts/project/readback
```

The default console operation verifies unchanged bytes, checks imported work/cost against the delivery fixture, edits its name/notes, and authors a separate project. The application oracle records tasks, resource types, assignments, calendar exceptions, producer build, and hashes. Optional `-ExpectedPath` accepts a JSON subset of expected records, selected by UID or name, and fails when selected values differ. `-NativeRoundTrip` saves another MPP through Microsoft Project, reopens it, and fails if the observed model changes. `-ApplicationAlerts` enables application dialogs for an attended run.

Other console operations:

| Command | Arguments | Purpose |
| --- | --- | --- |
| `corpus` | fixture directory, new output directory | Validate and edit every XML fixture |
| `edit-cases` | fixture directory, new output directory | Produce calendar date/removal and resource authoring cases with expected readback subsets |
| `schema` | input XML, local XSD | Validate offline; report whether the documented application namespace alias was applied |
| `cancellation` | 100,000-task input XML, new output directory | Request cancellation during load/save, check the five-second response budget and unchanged file destination |
| `native-probe` | paired MPP path, new output directory | Bounded compound/record inspection, stream-preserving rewrite, fixed-width name edit, and minimal new-container experiment |
| `native-corpus` | paired fixture path/directory, new output directory, optional file pattern | Compare native/MPX fields and effective calendars with independent XML; retain defaults, calculated values, cache differences, and unqualified fields as separate observations |
| `native-author-proof` | output MPP path, optional `Mpp8`, `Mpp9`, `Mpp12`, or `Mpp14` | Author a native project without a seed, with hierarchy, dependency, resource assignment, calendars, baseline, and custom alias/value |
| `native-lifecycle-proof` | source MPP path, new output directory | Exercise field growth, add/delete/reparent, calendar edits, templates, and template instantiation while retaining the source generation |
| `mpx-lifecycle-proof` | `new` or source MPX path, new output directory | Exercise unchanged save, field growth, structural/calendar edits, notes, and clearing values |
| `mpx-encoding-proof` | new output directory | Write representative Windows-1252, DOS 437/850, and Macintosh Roman fixtures |
| `mpx-scale` | task count (1–9000), `load` or `cancel` | Load 30 assignments and up to eight predecessor links per task; verify entity/link/work totals and report timing, allocation, or cancellation |
| `conversion-matrix` | new output directory | Assess, save, and reopen all 100 format pairs; verify core identities and mapped assignment values; record losses |
| `native-schema` | producer MPP path, new output directory | Reproduce generation-specific storage definitions without copying document records, strings, process pointers, or template payloads |
| `native-edit-proof` | producer `delivery.mpp` path, new output directory | Exercise field growth, structural changes, identities, calendar bindings, custom scalars, all baseline slots, templates, and XML conversion through the public API |
| `schedule-corpus` | paired fixture directory, new output directory | Compare native/XML calculations with producer dates, float, and critical flags |
| `schedule-scale` | task count, `dense` or `cancel` | Build a graph with up to four predecessors per task; verify its finish independently or exercise cooperative cancellation |
| `schedule-edit` | `authored` or XML path, new output directory | Author a calculated resource/calendar schedule, or exercise guarded recalculation of an imported input |
| `schedule-readback` | expected JSON, producer re-export XML | Verify applied dates after the independent application opens and exports calculated XML |
| `native-layout` / `native-records` | MPP path, new output directory | Inspect inert stream bytes or producer-mapped values for differential format research |
| `scale-create` | output XML, task count, shape | Generate one deterministic scale input |
| `scale-read-edit-save` | input XML, task count, shape | Exercise lifecycle and independently verify output records |

The native probe requires the paired XML beside the MPP for record comparison. Its minimal-container experiment is intentionally incomplete: conventional stream names alone are not a valid MPP document. Use `native-author-proof` and `native-edit-proof` for the public writer, then run their outputs through the application oracle with `-NativeRoundTrip`. The oracle's re-export XML exposes additional values for independent checks of baselines, custom fields, and calendar references. `New-ProjectFixtures.ps1 -ScenarioNames` also supports constraints, backward schedules, work/cost rules, local custom fields/all baseline slots, dated work weeks, and protected-file scenarios.

Use `New-ProjectFixtures.ps1 -NativeFormat MPP12` to create the tested Project 2007 export from a modern installed application. The normal output uses MPP14. `-NativeRoundTrip` saves in the installed application's default native format; use the recorded input/output generation when interpreting a legacy import check.

The optional [independent verifier](../../OfficeIMO.Project.IndependentVerification/README.md) uses MPXJ only as an external reader/writer oracle. It is outside the normal solution and runtime package. [Historical fixture provenance](historical-fixtures.json) records immutable source URLs, hashes, and the upstream repository license. Those third-party binaries are not redistributed here. Download only fixtures needed for a selected check into a separate verification directory and verify their SHA-256 values.

Independent-reader acceptance, mapped-field comparison, and Microsoft Project readback are separate checks. The matrix's core identity checks alone do not establish calendar, custom-field, or application fidelity. Run the independent export and `native-corpus` comparison, then the application oracle on formats that the installed application supports. Calendar comparisons include ordinary weeks and dated exceptions; observations distinguish stored values from independently derived defaults, WBS, remaining duration, and critical flags.

## Offline schema

Use the Microsoft Project 2013 SDK's `Documentation/Schemas/Project Client/mspdi_pj15.xsd`. Supply the local file path; the verifier does not download schemas or resolve imports over the network. The schema is Microsoft content and is not redistributed here. Its namespace differs from actual Project application exports, so the verifier reports an explicit namespace-only alias when needed. A successful check qualifies the selected authored subset, not every Project 2024 extension.

## Scale measurements

```powershell
Import-Module PSPublishModule
$result = Invoke-BenchmarkSuite -Path Build/Project/project-xml.benchmark.ps1
$result.Summary | Format-Table Scenario, Status, MedianMs, FailureCount
if ($result.Summary | Where-Object { $_.FailureCount -gt 0 -or $_.MedianMs -gt 30000 }) {
    throw 'A Project lifecycle validation or runtime budget failed.'
}
```

The suite uses one warmup and three measured iterations, rotated order, and no discarded outliers. Generation is excluded; process startup and independent readback are included. Allocation covers load/edit/save, while peak working set covers the whole child process. Correctness checks validate identities, assignment counts, edited values, relationships, and interval counts before results are accepted. The suite checks 2 GiB peak process memory and 4 GiB allocation limits; the invocation above checks the 30-second median runtime budget.

Record CPU topology, affinity, priority, power mode, SDK/runtime, input hashes, and the candidate revision with retained evidence. Do not compare timings from unrelated CPU/power configurations as if they were equivalent. Synthetic scale inputs do not represent all opaque structures in real MPP/XML exports.

`Invoke-BenchmarkSuite -Path Build/Project/project-schedule.benchmark.ps1` measures startup, graph creation, calculation, and independent validation for 1,000/10,000/100,000 tasks. Each one-minute task has up to four predecessors; a closed-form working-week finish and zero critical-path float validate the result. Its allocation metric covers calculation and its peak-memory metric covers the child process. Use `schedule-scale 100000 cancel` separately to check the five-second cancellation-response budget.

Application readback limitations, native outcomes, and runtime boundaries are listed in [the package support matrix](../../OfficeIMO.Project/SUPPORT.md). Keep compact reproducible evidence and fixture manifests; remove superseded outputs, SDK extraction files, and scale inputs after validation.
