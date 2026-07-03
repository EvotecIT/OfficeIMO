# OfficeIMO CI and Test Strategy

This note records the current cleanup direction for OfficeIMO tests and build warnings. It is meant for maintainers working on CI, test placement, or warning policy.

## Current Shape

The solution has a legacy aggregate test project, `OfficeIMO.Tests`, plus domain projects such as `OfficeIMO.Excel.Tests`, `OfficeIMO.CSV.Tests`, `OfficeIMO.VerifyTests`, and `OfficeIMO.MarkdownRenderer.Wpf.Tests`.

`OfficeIMO.Tests` still covers many product areas: PDF, Word, Markdown, Visio, RTF, PowerPoint, HTML, Reader, Drawing, Markup, and shared workflow tests. That made sense while features were being built quickly, but it is now too large for clean CI ownership. A full rebuild of the aggregate project used to emit thousands of duplicate nullable-warning lines across target frameworks, which made GitHub annotations noisy and hid the warnings that matter.

`OfficeIMO.Excel.Tests` is the first real split from the aggregate. It owns the Excel test sources, Excel image export, Excel PDF, Google Sheets payload, and Excel compatibility contracts. The project has its own references and friend-assembly access, so Excel test warnings and internal contracts no longer ride through the whole aggregate test assembly.

## Decision

It is time to split tests by product/domain project instead of continuing to grow the aggregate project.

The desired end state is:

- `OfficeIMO.Word.Tests` for Word and Word conversion contracts.
- `OfficeIMO.Excel.Tests` for Excel, Excel image export, Excel PDF, and Excel compatibility contracts. This project exists now and should be the target for new Excel tests.
- `OfficeIMO.Pdf.Tests` for native PDF contracts and PDF compliance/readback tests.
- `OfficeIMO.Markdown.Tests` for Markdown parsing, rendering, Markdig parity, and Markdown conversion contracts.
- `OfficeIMO.Visio.Tests`, `OfficeIMO.Rtf.Tests`, `OfficeIMO.PowerPoint.Tests`, `OfficeIMO.Html.Tests`, and `OfficeIMO.Reader.Tests` for their domain contracts.
- Small integration or workflow projects only when a test intentionally crosses several domains.

Do not split tests only by folder size. Split when the test project can own a real contract, build a smaller dependency graph, and run independently in CI.

## CI Direction

The Ubuntu test lane runs named partitions in a bounded matrix. This keeps the existing contract coverage but avoids one long serial job that can run close to the hosted-runner timeout.

The current partitions are:

- `markdown-large`
- `markdown-suite`
- `pdf-visual-inspector`
- `pdf-core`
- `excel-image-charts`
- `excel-legacy-reader`
- `excel-core-named`
- `word-rtf-html`
- `other-projects`

Excel partitions run `OfficeIMO.Excel.Tests` directly. The remaining partitions continue to run `OfficeIMO.Tests` until their domains are split into their own projects.

Keep `max-parallel` bounded so the workflow improves wall-clock time without flooding the organization with too many simultaneous jobs.

Coverage is not collected in the PR matrix. Data-heavy Markdown and PDF tests become much slower under coverage instrumentation, so coverage should move to a separate scheduled or manually dispatched lane if the project needs it.

For test jobs, prefer building the test project for the target framework instead of rebuilding the whole solution. The cross-platform build job already proves the solution build. Test jobs should prove test contracts and keep their dependency graph as small as practical.

## Warning Policy

Production projects keep warnings as errors.

The legacy aggregate `OfficeIMO.Tests` project suppresses nullable warnings, platform/framework analyzer warnings, and a few xUnit style analyzer warnings while the suite is being split, because the current volume makes CI annotations unusable. `OfficeIMO.Excel.Tests` also carries a scoped transitional suppression for the existing moved Excel tests so the real project split can land without reintroducing annotation spam. New Excel tests should avoid adding new nullable debt, and follow-up cleanup should remove suppressions as the moved tests are made nullable-clean.

New domain test projects should start clean:

- Nullable enabled.
- No blanket nullable `NoWarn`.
- Warnings fixed when they point to unclear test setup or real null contracts.
- Intentional null inputs expressed with nullable types, null-forgiving operators, or focused helper methods where the test contract requires them.

Example projects are executable documentation. Prefer fixing example warnings over suppressing them, because users copy those files as usage guidance.

## Useful Local Commands

Restore and build the current aggregate test project:

```powershell
dotnet restore OfficeIMO.Tests/OfficeIMO.Tests.csproj
dotnet build OfficeIMO.Tests/OfficeIMO.Tests.csproj --configuration Release --framework net8.0 --no-restore
```

Restore and build the Excel test project:

```powershell
dotnet restore OfficeIMO.Excel.Tests/OfficeIMO.Excel.Tests.csproj
dotnet build OfficeIMO.Excel.Tests/OfficeIMO.Excel.Tests.csproj --configuration Release --framework net8.0 --no-restore
```

Run a focused Excel partition locally:

```powershell
dotnet test OfficeIMO.Excel.Tests/OfficeIMO.Excel.Tests.csproj --configuration Release --framework net8.0 --filter "FullyQualifiedName~OfficeIMO.Tests.ExcelImageExport" --logger "console;verbosity=minimal"
```

Run the dedicated Excel image visual smoke gate:

```powershell
./Build/Test-ExcelImageVisualGate.ps1 -Configuration Release -Framework net8.0 -Suite Smoke
```
