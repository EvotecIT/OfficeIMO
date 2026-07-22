---
title: NativeAOT Deployment
description: Publish OfficeIMO applications as native executables, choose AOT-safe APIs, and understand optional integration boundaries.
order: 80
---

OfficeIMO's NativeAOT evidence covers the complete production project inventory rather than a hand-picked package list. The repository currently has **89 production projects**: **88 publish and execute in NativeAOT validation**, while the WPF/WebView2 renderer is deliberately tested as a managed Windows component because the .NET SDK rejects trimming for WPF executables (`NETSDK1168`).

The 88 native-validated projects are not all proved in the same way:

- **85 production libraries** are retained as complete assemblies in one native compile graph; the resulting executable must start successfully.
- **1 optional Google APIs adapter** runs a bounded token-store workflow natively. Its complete Google authorization dependency surface is not advertised as trim-safe because fully rooting `Google.Apis` and `Newtonsoft.Json` produces upstream warnings.
- **2 production command-line tools** publish as native executables and must start and return their real command help.

The [machine-readable project matrix](/data/aot-compatibility.json) names all 89 projects and records which proof applies to each one. This distinction matters to customers: a green native workflow is useful evidence, but it is not permission to assume that every optional third-party API has been executed.

## Publish your application

Use .NET 8 or .NET 10 and enable NativeAOT in the application that consumes OfficeIMO:

```xml
<PropertyGroup>
  <TargetFramework>net10.0</TargetFramework>
  <PublishAot>true</PublishAot>
</PropertyGroup>
```

Publish for the operating system and architecture you plan to deploy:

```powershell
dotnet publish -c Release -r win-x64
dotnet publish -c Release -r linux-x64
```

NativeAOT evaluates your complete application, including every package and code path you use. Run the published executable and assert the generated or extracted document content before shipping it.

The Word, Excel, PowerPoint, and Word-to-HTML packages use the Microsoft Open XML SDK in the supported range `[3.5.1, 4.0.0)`. Normal OfficeIMO consumers receive that dependency transitively. If your application references `DocumentFormat.OpenXml` directly, keep its resolved version inside the same range so the dependency graph matches the versions validated by OfficeIMO.

## Project-level coverage

| Production classification | Projects | What CI proves | Customer guidance |
|---|---:|---|---|
| Fully rooted libraries | 85 | The complete assembly surfaces compile into one NativeAOT executable on Windows and Linux, and that executable starts. | These packages are suitable NativeAOT building blocks; still test the exact documents and options your application uses. |
| Bounded Google APIs adapter | 1 | `OfficeIMO.GoogleWorkspace.Auth.GoogleApis` constructs its data-store adapter and round-trips a value in the native executable. | The validated adapter path is native. Treat live OAuth/provider flows as application-specific until your chosen Google dependency graph publishes cleanly. |
| Native command-line tools | 2 | `OfficeIMO.Markup.Cli` and `OfficeIMO.Reader.Tool` publish and start as native executables on Windows and Linux. | Native CLI deployment is supported; validate the concrete commands and formats used by your job. |
| Managed Windows UI | 1 | `OfficeIMO.MarkdownRenderer.Wpf` builds and runs through the managed Windows/WPF test lane. | Do not enable NativeAOT for this WPF/WebView2 UI package. Use the managed Windows deployment model. |

CI fails if a production project is added, removed, or renamed without being classified in this matrix.

## Executed document workflows

The repository publishes separate native applications so one optional integration cannot hide the status of an unrelated document workflow.

| Product area | `win-x64` | `linux-x64` | Native executable verifies |
|---|---|---|---|
| Word | Pass | Pass | Create a DOCX, save it, reopen it, and read the expected paragraph. |
| Excel | Pass | Pass | Create a typed table, save it, reopen it, and read typed tabular data. |
| PowerPoint | Pass | Pass | Create a chart and its embedded workbook, duplicate the slide and relationships, save, and reopen both charts. |
| Markdown | Pass | Pass | Compose and render a document through the fluent API. |
| CSV | Pass | Pass | Parse a document and inspect its schema. |
| Reader CSV | Pass | Pass | Route CSV through the focused adapter and read normalized table chunks. |
| Reader complete preset | Pass | Pass | Register all 30 local in-process handlers and perform representative structured extraction. |
| HTML, PDF, and images | Pass | Pass | Render HTML to SVG and PNG, create a searchable PDF, and read marker text back. |

These eight document workflows complement the project-level compile matrix with behavior and output checks. They are useful end-to-end baselines, not a claim that every possible document, option, or third-party service has been exercised. Add your templates, accepted formats, fonts, and representative source files to your application tests.

## Package-family guidance

| Package family | NativeAOT guidance |
|---|---|
| `OfficeIMO.Word`, `OfficeIMO.Excel`, `OfficeIMO.PowerPoint` | Use the normal typed document APIs. The common create, edit, relationship, save, and reload paths are covered by native executables. |
| `OfficeIMO.Markdown`, `OfficeIMO.CSV`, `OfficeIMO.Html`, `OfficeIMO.Pdf` | In-process composition, parsing, and rendering are AOT-friendly. Validate output fidelity with your real content. |
| `OfficeIMO.Reader.*` | The `OfficeIMO.Reader.All` preset registers all local format handlers in NativeAOT. Add only the adapters you need when binary size matters. |
| Format and conversion adapters | The complete production library assemblies are rooted in the native compile graph. Test the exact conversion direction and fidelity your application accepts. |
| Google Workspace, Confluence, and other network clients | The dependency-light OfficeIMO client libraries are fully rooted in the native matrix. The optional `Google.Apis` credential adapter has a bounded native token-store test; publish and test the live authentication and HTTP provider selected by your application. |
| OCR process and Tesseract adapters | The OfficeIMO host adapter can be native; OCR still runs in the configured external executable and must be deployed separately. |
| WPF/WebView2 renderer | Use managed Windows deployment. The .NET SDK currently rejects trimmed WPF executables with `NETSDK1168`, so OfficeIMO does not market this package as NativeAOT-compatible. |

## Prefer typed data paths

NativeAOT works best when types are visible to the compiler. OfficeIMO's normal document APIs already follow that model. APIs that inspect an arbitrary runtime object graph are marked so the compiler can warn at the call site.

Recommended patterns include:

- write Excel tables from `DataTable`, explicit cell values, or typed row mappings;
- read Excel rows with generic typed readers whose model type is known at publish time;
- build Word tables from explicit columns and cells when the model shape is dynamic;
- use generic Markdown object/table builders or dictionary-based data when fields are selected at runtime;
- register Reader adapters explicitly, or use `AddAllOfficeIMOHandlers()` for the complete local preset.

If your application loads plug-ins, type names, templates, or model members dynamically, preserve those members in the application or replace discovery with an explicit mapping. This is an application boundary rather than an OfficeIMO-specific switch.

## Validate the deployment you will ship

A practical NativeAOT acceptance test should:

1. publish for the target runtime identifier;
2. start the produced native executable;
3. create, convert, or read a representative document;
4. reopen the output where the format supports it;
5. assert useful content such as text, tables, formulas, slides, relationships, or searchable PDF text.

OfficeIMO contributors can run the checked-in matrix for the current machine:

```powershell
./Build/Test-AotScenarios.ps1
```

The script first validates the production-library host, then runs the eight document workflows and the two production CLI startup checks. Each scenario receives a fresh SDK artifacts directory, so a Linux run cannot consume Windows `obj` metadata and one scenario cannot reuse another scenario's native state. The same matrix runs in repository CI on Windows and Linux. Optional JSON results can be retained for deployment evidence:

```powershell
./Build/Test-AotScenarios.ps1 -JsonOutputPath ./artifacts/aot-results.json
./Build/Test-AotProjectCoverage.ps1 -JsonOutputPath ./artifacts/aot-project-matrix.json
```

## Trimming, ReadyToRun, and single-file deployment

These deployment modes solve different problems:

| Goal | Setting |
|---|---|
| Native executable and fast startup | `<PublishAot>true</PublishAot>` |
| Smaller managed deployment | `<PublishTrimmed>true</PublishTrimmed>` |
| Faster managed startup with broader runtime compatibility | `<PublishReadyToRun>true</PublishReadyToRun>` |
| One managed deployment file | `<PublishSingleFile>true</PublishSingleFile>` |

Do not copy warning suppressions from another application. Treat every trim or AOT warning at your call site as a request to choose a typed API, preserve an intentionally dynamic model, or remove a dependency that cannot support the deployment target.
