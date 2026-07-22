---
title: AOT and Trimming
description: Executed NativeAOT scenarios, current compiler blockers, and a reproducible validation path for OfficeIMO packages.
order: 80
---

OfficeIMO does not advertise one blanket NativeAOT result for its entire package graph. Compatibility is recorded for a concrete publish target and an executable workflow: the project must restore, compile to native code, start, produce its output, and read key data back where the format supports a round trip.

## Current executable matrix

The pass rows below were published and executed as .NET 10 `win-x64` native binaries. The repository CI repeats the supported matrix on `linux-x64`; it fails if any passing scenario stops compiling or running. Excel and PowerPoint have separate smoke projects so their current compiler blockers cannot hide the passing Word, Markdown, CSV, Reader, or rendering paths.

| Scenario | `win-x64` result | What the native binary proves |
|---|---|---|
| Word | Pass | Creates a DOCX, saves it, reloads it, and finds the marker paragraph. |
| Markdown | Pass | Builds and renders Markdown through the fluent document API. |
| CSV | Pass | Parses CSV and reads the resulting header schema. |
| Reader CSV | Pass | Routes CSV through an isolated Reader handler and emits normalized table chunks. |
| HTML, PDF, and image rendering | Pass | Parses HTML, renders SVG and PNG, creates PDF, and extracts searchable marker text. |
| Excel | Publish blocked | NativeAOT reports `IL2072` in `DirectDataSetTableModel.ToDataTable`; the round-trip binary is intentionally retained for rechecks. |
| PowerPoint | Publish blocked | NativeAOT reports `IL2060`, `IL2075`, `IL2087`, and `IL3050` around reflective Open XML part creation and media relationships. |

This table does **not** turn one passing workflow into a promise for every API in that assembly. Packages and adapters not listed here are **not tested by this matrix**.

## Reproduce the matrix

Run all supported publish-and-execute scenarios for the current machine:

```powershell
./Build/Test-AotScenarios.ps1
```

Recheck the known Excel and PowerPoint blockers too:

```powershell
./Build/Test-AotScenarios.ps1 -IncludeKnownBlocked
```

The script uses one NativeAOT executable per scenario, verifies the expected status, and deletes its temporary publish output. Pass an explicit runtime identifier when validating a deployment target:

```powershell
./Build/Test-AotScenarios.ps1 -RuntimeIdentifier linux-x64
```

## Why separate executables matter

NativeAOT analyzes the complete dependency graph of an application. If Excel and Markdown share one smoke executable, a reflection blocker in Excel prevents the binary from being produced and tells you nothing about Markdown. The checked-in smoke projects isolate the public surface being evaluated while still executing a meaningful workflow rather than an empty startup check.

## Trimming analyzers are a different signal

The CI job also builds the solution for .NET 8 and .NET 10 with `EnableTrimAnalyzer` and `EnableAotAnalyzer`. Analyzer success is useful source evidence, but it is weaker than publishing and executing a native binary. Likewise, suppressing a warning or rooting an assembly can preserve code without proving the resulting workflow is correct.

For a trimmed deployment that is not NativeAOT, start conservatively and test the same read/write behavior your service will use:

```xml
<PropertyGroup>
  <PublishTrimmed>true</PublishTrimmed>
  <TrimMode>link</TrimMode>
</PropertyGroup>
```

Do not copy warning suppressions from another application without confirming that the reflected members and file paths it needs are preserved.

## Choosing a deployment mode

| Requirement | Starting point |
|---|---|
| Small native CLI using a proven path | Use one of the passing smoke projects as the minimum consumer and add your exact workflow. |
| Office document workflow not represented above | Fork the closest smoke project and prove create, save, reopen, and semantic readback. |
| Excel or PowerPoint NativeAOT today | Treat native publication as blocked until the listed compiler diagnostics are resolved and the retained smoke executable runs. |
| Faster startup without full AOT | Evaluate ReadyToRun with `<PublishReadyToRun>true</PublishReadyToRun>`. |
| One deployable file | Evaluate self-contained single-file publishing separately; it is not equivalent to NativeAOT. |

Framework support, COM-free execution, trimming-analyzer cleanliness, NativeAOT publication, and rendering fidelity are separate contracts. Record each one at the level you actually validated.
