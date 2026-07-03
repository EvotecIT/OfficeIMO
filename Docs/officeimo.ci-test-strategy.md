# OfficeIMO CI and Test Strategy

This note records the current cleanup direction for OfficeIMO tests and build warnings. It is meant for maintainers working on CI, test placement, or warning policy.

## Current Shape

The solution has a legacy aggregate test project, `OfficeIMO.Tests`, plus domain projects such as `OfficeIMO.Word.Tests`, `OfficeIMO.Excel.Tests`, `OfficeIMO.PowerPoint.Tests`, `OfficeIMO.Visio.Tests`, `OfficeIMO.Rtf.Tests`, `OfficeIMO.Html.Tests`, `OfficeIMO.Reader.Tests`, `OfficeIMO.Pdf.Tests`, `OfficeIMO.Markdown.Tests`, `OfficeIMO.Drawing.Tests`, `OfficeIMO.Markup.Tests`, `OfficeIMO.CSV.Tests`, `OfficeIMO.VerifyTests`, and `OfficeIMO.MarkdownRenderer.Wpf.Tests`.

`OfficeIMO.Tests` now covers shared workflow and guardrail tests that intentionally span product areas. That made sense while features were being built quickly, but it should no longer own product-domain suites. A full rebuild of the aggregate project used to emit thousands of duplicate nullable-warning lines across target frameworks, which made GitHub annotations noisy and hid the warnings that matter.

`OfficeIMO.Word.Tests`, `OfficeIMO.Excel.Tests`, `OfficeIMO.PowerPoint.Tests`, `OfficeIMO.Visio.Tests`, `OfficeIMO.Rtf.Tests`, `OfficeIMO.Html.Tests`, `OfficeIMO.Reader.Tests`, `OfficeIMO.Pdf.Tests`, `OfficeIMO.Markdown.Tests`, `OfficeIMO.Drawing.Tests`, and `OfficeIMO.Markup.Tests` are real splits from the aggregate. `OfficeIMO.Word.Tests` owns the core Word API, Word PDF conversion, Word/Markdown round-trip conversion tests, Google Docs payload, and fixture contracts. `OfficeIMO.Excel.Tests` owns the Excel test sources, Excel image export, Excel PDF, Google Sheets payload, and Excel compatibility contracts. `OfficeIMO.PowerPoint.Tests` owns the PowerPoint presentation API, PowerPoint image export, and PowerPoint PDF conversion tests. `OfficeIMO.Visio.Tests` owns the Visio document, stencil, diagram, SVG/PNG export, validation, and visual baseline contracts. `OfficeIMO.Rtf.Tests` owns native RTF, Word RTF conversion, RTF Markdown, RTF PDF, lossless editor, tokenizer, syntax, and golden corpus contracts. `OfficeIMO.Html.Tests` owns shared HTML policy, HTML ingestion, HTML PDF, and Office-to-HTML bridge contracts. `OfficeIMO.Reader.Tests` owns the unified reader API, modular reader adapters, reader packaging guardrails, and golden reader fixtures. `OfficeIMO.Pdf.Tests` owns native PDF, PDF compliance/readback, PDF visual-quality, raster baseline, and PDF conversion scenario contracts. `OfficeIMO.Markdown.Tests` owns pure Markdown parsing, rendering, Markdig/CommonMark/GFM parity, native/source-map behavior, HTML-to-Markdown, transcript rendering, golden fixtures, and Markdown conversion contracts that belong with the Markdown engine. `OfficeIMO.Drawing.Tests` owns shared Drawing primitives, raster/SVG rendering, image composition, sparkline rendering, shape presets, and architecture guardrails for Drawing consumers. `OfficeIMO.Markup.Tests` owns Office Markup parsing, emitters, and Word/Excel/PowerPoint export contracts for the Markup authoring surface. These projects have their own references and friend-assembly access, so their test warnings and internal contracts no longer ride through the whole aggregate test assembly.

## Decision

It is time to split tests by product/domain project instead of continuing to grow the aggregate project.

The desired end state is:

- `OfficeIMO.Word.Tests` for core Word and Word conversion contracts. This project exists now and should be the target for new core Word and Word PDF conversion tests.
- `OfficeIMO.Excel.Tests` for Excel, Excel image export, Excel PDF, and Excel compatibility contracts. This project exists now and should be the target for new Excel tests.
- `OfficeIMO.Pdf.Tests` for native PDF, PDF compliance/readback, visual-quality, raster baseline, and PDF conversion scenario contracts. This project exists now and should be the target for new PDF tests.
- `OfficeIMO.Markdown.Tests` for pure Markdown parsing, rendering, Markdig/CommonMark/GFM parity, Markdown HTML conversion, source-map/native behavior, transcript rendering, golden fixtures, and Markdown conversion contracts. This project exists now and should be the target for new Markdown tests.
- `OfficeIMO.Visio.Tests` for Visio document, stencil, diagram, export, validation, and visual baseline contracts. This project exists now and should be the target for new Visio tests.
- `OfficeIMO.Rtf.Tests` for native RTF, Word RTF conversion, RTF Markdown, RTF PDF, lossless editor, tokenizer, syntax, and golden corpus contracts. This project exists now and should be the target for new RTF tests.
- `OfficeIMO.Html.Tests` for shared HTML policy, HTML ingestion, HTML PDF, and Office-to-HTML bridge contracts. This project exists now and should be the target for new HTML tests.
- `OfficeIMO.Reader.Tests` for unified reader, modular reader, packaging, and golden fixture contracts. This project exists now and should be the target for new Reader tests.
- `OfficeIMO.Drawing.Tests` for Drawing primitives, raster/SVG rendering, image composition, sparkline rendering, shape presets, and Drawing architecture guardrails. This project exists now and should be the target for new Drawing tests.
- `OfficeIMO.Markup.Tests` for Markup parsing, emitters, and Word/Excel/PowerPoint export contracts. This project exists now and should be the target for new Markup tests.
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

Word, Excel, RTF, HTML, Reader, PDF, Markdown, Drawing, and Markup partitions run `OfficeIMO.Word.Tests`, `OfficeIMO.Excel.Tests`, `OfficeIMO.Rtf.Tests`, `OfficeIMO.Html.Tests`, `OfficeIMO.Reader.Tests`, `OfficeIMO.Pdf.Tests`, `OfficeIMO.Markdown.Tests`, `OfficeIMO.Drawing.Tests`, and `OfficeIMO.Markup.Tests` directly. The `other-projects` partition runs `OfficeIMO.Drawing.Tests`, `OfficeIMO.Markup.Tests`, `OfficeIMO.PowerPoint.Tests`, `OfficeIMO.Visio.Tests`, and `OfficeIMO.Reader.Tests` directly before the remaining aggregate workflow/guardrail filters. The `markdown-large` partition also runs the Word/Markdown round-trip tests from `OfficeIMO.Word.Tests`, because those tests exercise Word document conversion behavior even though their class names are Markdown-oriented. The remaining aggregate filters continue to run `OfficeIMO.Tests` until those shared tests are moved to a more precise owner or intentionally kept as cross-domain guardrails.

Keep `max-parallel` bounded so the workflow improves wall-clock time without flooding the organization with too many simultaneous jobs.

Coverage is not collected in the PR matrix. Data-heavy Markdown and PDF tests become much slower under coverage instrumentation, so coverage should move to a separate scheduled or manually dispatched lane if the project needs it.

For test jobs, prefer building the test project for the target framework instead of rebuilding the whole solution. The cross-platform build job already proves the solution build. Test jobs should prove test contracts and keep their dependency graph as small as practical.

## Warning Policy

Production projects keep warnings as errors.

The legacy aggregate `OfficeIMO.Tests` project suppresses nullable warnings, platform/framework analyzer warnings, and a few xUnit style analyzer warnings while the suite is being split, because the current volume makes CI annotations unusable. `OfficeIMO.Word.Tests`, `OfficeIMO.Excel.Tests`, `OfficeIMO.PowerPoint.Tests`, `OfficeIMO.Visio.Tests`, `OfficeIMO.Rtf.Tests`, `OfficeIMO.Html.Tests`, `OfficeIMO.Reader.Tests`, `OfficeIMO.Pdf.Tests`, `OfficeIMO.Markdown.Tests`, `OfficeIMO.Drawing.Tests`, and `OfficeIMO.Markup.Tests` also carry scoped transitional suppressions for the existing moved tests so the real project split can continue without reintroducing annotation spam. New Word, Excel, PowerPoint, Visio, RTF, HTML, Reader, PDF, Markdown, Drawing, and Markup tests should avoid adding new nullable debt, and follow-up cleanup should remove suppressions as the moved tests are made nullable-clean.

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

Restore and build the Word test project:

```powershell
dotnet restore OfficeIMO.Word.Tests/OfficeIMO.Word.Tests.csproj
dotnet build OfficeIMO.Word.Tests/OfficeIMO.Word.Tests.csproj --configuration Release --framework net8.0 --no-restore
```

Restore and build the Excel test project:

```powershell
dotnet restore OfficeIMO.Excel.Tests/OfficeIMO.Excel.Tests.csproj
dotnet build OfficeIMO.Excel.Tests/OfficeIMO.Excel.Tests.csproj --configuration Release --framework net8.0 --no-restore
```

Restore and build the Visio test project:

```powershell
dotnet restore OfficeIMO.Visio.Tests/OfficeIMO.Visio.Tests.csproj
dotnet build OfficeIMO.Visio.Tests/OfficeIMO.Visio.Tests.csproj --configuration Release --framework net8.0 --no-restore
```

Restore and build the RTF test project:

```powershell
dotnet restore OfficeIMO.Rtf.Tests/OfficeIMO.Rtf.Tests.csproj
dotnet build OfficeIMO.Rtf.Tests/OfficeIMO.Rtf.Tests.csproj --configuration Release --framework net8.0 --no-restore
```

Restore and build the HTML test project:

```powershell
dotnet restore OfficeIMO.Html.Tests/OfficeIMO.Html.Tests.csproj
dotnet build OfficeIMO.Html.Tests/OfficeIMO.Html.Tests.csproj --configuration Release --framework net8.0 --no-restore
```

Restore and build the Reader test project:

```powershell
dotnet restore OfficeIMO.Reader.Tests/OfficeIMO.Reader.Tests.csproj
dotnet build OfficeIMO.Reader.Tests/OfficeIMO.Reader.Tests.csproj --configuration Release --framework net8.0 --no-restore
```

Restore and build the PDF test project:

```powershell
dotnet restore OfficeIMO.Pdf.Tests/OfficeIMO.Pdf.Tests.csproj
dotnet build OfficeIMO.Pdf.Tests/OfficeIMO.Pdf.Tests.csproj --configuration Release --framework net8.0 --no-restore
```

Restore and build the Markdown test project:

```powershell
dotnet restore OfficeIMO.Markdown.Tests/OfficeIMO.Markdown.Tests.csproj
dotnet build OfficeIMO.Markdown.Tests/OfficeIMO.Markdown.Tests.csproj --configuration Release --framework net8.0 --no-restore
```

Restore and build the Drawing test project:

```powershell
dotnet restore OfficeIMO.Drawing.Tests/OfficeIMO.Drawing.Tests.csproj
dotnet build OfficeIMO.Drawing.Tests/OfficeIMO.Drawing.Tests.csproj --configuration Release --framework net8.0 --no-restore
```

Restore and build the Markup test project:

```powershell
dotnet restore OfficeIMO.Markup.Tests/OfficeIMO.Markup.Tests.csproj
dotnet build OfficeIMO.Markup.Tests/OfficeIMO.Markup.Tests.csproj --configuration Release --framework net8.0 --no-restore
```

Run a focused Excel partition locally:

```powershell
dotnet test OfficeIMO.Excel.Tests/OfficeIMO.Excel.Tests.csproj --configuration Release --framework net8.0 --filter "FullyQualifiedName~OfficeIMO.Tests.ExcelImageExport" --logger "console;verbosity=minimal"
```

Run the dedicated Excel image visual smoke gate:

```powershell
./Build/Test-ExcelImageVisualGate.ps1 -Configuration Release -Framework net8.0 -Suite Smoke
```
