---
title: "AOT and Trimming: Current State in OfficeIMO"
description: "The executable NativeAOT matrix for OfficeIMO, including passing document workflows, current compiler blockers, and how to reproduce both."
date: 2025-11-01
tags: [aot, trimming, performance]
author: OfficeIMO Team
---

.NET NativeAOT is not a label that can be assigned from a project file or a short dependency list. OfficeIMO contains focused document engines, format adapters, readers, and renderers, so each useful claim needs a concrete application graph and an executed workflow.

## The result we can prove

The repository now keeps one native smoke executable per evaluated workflow. On `win-x64`, the following .NET 10 binaries publish and run:

- Word creates a DOCX, saves it, reloads it, and finds a marker paragraph.
- Markdown composes and renders a fluent document.
- CSV parses data and exposes the expected header schema.
- Reader routes CSV through an isolated handler and emits normalized table chunks.
- HTML rendering produces SVG, PNG, and searchable PDF output, then reads the marker text back from the PDF.

CI repeats the supported matrix on `linux-x64`. A passing row therefore means more than “the compiler accepted an empty program,” but less than “every member in every related package is safe.”

## The blockers are part of the story

Excel and PowerPoint have their own smoke projects and are deliberately retained even though native publication currently stops before runtime:

| Package | Current publish result | Compiler evidence |
|---|---|---|
| OfficeIMO.Excel | Blocked | `IL2072` in `DirectDataSetTableModel.ToDataTable` |
| OfficeIMO.PowerPoint | Blocked | `IL2060`, `IL2075`, `IL2087`, and `IL3050` in reflective Open XML part and media paths |

Keeping these as separate executables matters. If every package shared one test app, the first compiler error would hide which independent graphs already work. It would also make a future fix difficult to attribute.

## Run it yourself

The repository script publishes, executes, and cleans the supported matrix for the current machine:

```powershell
./Build/Test-AotScenarios.ps1
```

Add `-IncludeKnownBlocked` to confirm that the documented Excel and PowerPoint diagnostics still reproduce. The script fails if a passing workflow regresses or if a known blocked status changes without the compatibility contract being updated.

```powershell
./Build/Test-AotScenarios.ps1 -IncludeKnownBlocked
```

For a deployment target different from the host, pass its runtime identifier explicitly. Publishing alone is not enough when the output cannot run on the current machine, so the CI job remains the execution evidence for Linux.

## Trimming is related, not identical

The solution also builds with trim and AOT analyzers for .NET 8 and .NET 10. Those analyzers catch source patterns that may be unsafe, but analyzer cleanliness cannot replace a real native publish and semantic runtime assertion. Conversely, preserving an entire assembly can silence trimming failures while increasing output size and still leave an untested workflow.

The useful rule is straightforward: start from the smallest passing smoke that matches your package, add the exact APIs and documents your service uses, publish for the real runtime identifier, and assert the generated content. Anything outside that graph remains not tested until it gets its own evidence.

See the [AOT and trimming guide](/docs/advanced/aot-trimming/) for the live matrix and deployment choices.
