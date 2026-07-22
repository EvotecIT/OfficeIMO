---
title: NativeAOT Deployment
description: Publish OfficeIMO applications as native executables, choose AOT-safe APIs, and understand optional integration boundaries.
order: 80
---

OfficeIMO supports NativeAOT for its standard in-process document workflows. Word, Excel, PowerPoint, Markdown, CSV, the local-format Reader preset, and HTML/PDF/image rendering are published and executed as native test applications on Windows and Linux. Every production project is also built with the .NET trimming and AOT analyzers so unsafe dynamic paths must be explicit.

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

## Tested OfficeIMO workflows

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

These are useful end-to-end baselines, not a claim that every possible document or third-party service has been exercised. Add your templates, accepted formats, fonts, and representative source files to your application tests.

## Package-family guidance

| Package family | NativeAOT guidance |
|---|---|
| `OfficeIMO.Word`, `OfficeIMO.Excel`, `OfficeIMO.PowerPoint` | Use the normal typed document APIs. The common create, edit, relationship, save, and reload paths are covered by native executables. |
| `OfficeIMO.Markdown`, `OfficeIMO.CSV`, `OfficeIMO.Html`, `OfficeIMO.Pdf` | In-process composition, parsing, and rendering are AOT-friendly. Validate output fidelity with your real content. |
| `OfficeIMO.Reader.*` | The `OfficeIMO.Reader.All` preset registers all local format handlers in NativeAOT. Add only the adapters you need when binary size matters. |
| Format and conversion adapters | The production projects are analyzer-gated and use the same in-process engines. Test the exact conversion direction and fidelity your application accepts. |
| Google Workspace, Confluence, and other network clients | The OfficeIMO client layer can be used from an AOT host, but authentication providers, HTTP handlers, and live service behavior remain part of your application graph. Publish and test the provider you select. |
| OCR process and Tesseract adapters | The OfficeIMO host adapter can be native; OCR still runs in the configured external executable and must be deployed separately. |
| WPF/WebView2 renderer | This is a desktop UI integration rather than a NativeAOT document engine. Use the deployment mode supported by the selected WPF and WebView2 runtime. |

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

The script gives each scenario a fresh SDK artifacts directory, so a Linux run cannot consume Windows `obj` metadata and one scenario cannot reuse another scenario's native state. The same scenarios run in repository CI on Linux. An optional JSON result can be retained for deployment evidence:

```powershell
./Build/Test-AotScenarios.ps1 -JsonOutputPath ./artifacts/aot-results.json
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
