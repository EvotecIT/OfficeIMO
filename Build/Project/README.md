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

The default console operation verifies unchanged bytes, checks imported work/cost against the delivery fixture, edits its name/notes, and authors a separate project. The application oracle records tasks, resource types, assignments, calendar exceptions, producer build, and hashes. Optional `-ExpectedPath` accepts a JSON subset of expected records, selected by UID or name, and fails when selected values differ.

Other console operations:

| Command | Arguments | Purpose |
| --- | --- | --- |
| `corpus` | fixture directory, new output directory | Validate and edit every XML fixture |
| `edit-cases` | fixture directory, new output directory | Produce calendar date/removal and resource authoring cases with expected readback subsets |
| `schema` | input XML, local XSD | Validate offline; report whether the documented application namespace alias was applied |
| `native-probe` | paired MPP path, new output directory | Bounded compound/record inspection, stream-preserving rewrite, fixed-width name edit, and minimal new-container experiment |
| `scale-create` | output XML, task count, shape | Generate one deterministic scale input |
| `scale-read-edit-save` | input XML, task count, shape | Exercise lifecycle and independently verify output records |

The native probe requires the paired XML beside the MPP for record comparison. Run its outputs through the application oracle separately. The minimal native creation experiment is expected to fail: conventional stream names alone are not a valid MPP document. These are feasibility experiments, not a supported native codec.

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

Application readback limitations, native outcomes, and runtime boundaries are listed in [the package support matrix](../../OfficeIMO.Project/SUPPORT.md). Keep compact reproducible evidence and fixture manifests; remove superseded outputs, SDK extraction files, and scale inputs after validation.
