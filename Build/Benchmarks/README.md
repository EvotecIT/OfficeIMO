# Assistant evidence measurements

Build `OfficeIMO.Studio` in Release, then run the evidence suite from a fresh PowerShell 7 process:

```powershell
./Build/Benchmarks/Run-AssistantEvidenceBenchmarks.ps1
```

The PowerForge suite compares reading a PDF again with inspecting an already prepared immutable snapshot. Both paths validate the same source hash, evidence hash, page count and text coverage. The default matrix uses synthetic 1-, 100- and 500-page PDFs, two warmups and seven measured iterations per case. `-Pages`, `-WarmupCount`, `-IterationCount` and `-OutputRoot` select a different run; `-Plan` prints the execution plan.

Run from an unchanged checkout. PowerForge records source provenance and rejects runs if source changes during measurement. Keep other builds and benchmarks idle. On machines with multiple CPU or cache domains, use an explicitly recorded affinity and repeat on each relevant group.

Outputs include samples, summaries, comparison tables and environment metadata under the selected output folder. `OperationAllocatedBytes` measures managed allocations across the operation; it is not peak or retained memory. Elapsed time includes harness overhead, which matters for very short reuse operations. These measurements cover local evidence preparation and readiness only. They do not measure inference latency, model accuracy, OCR, or end-to-end UI responsiveness.

The separate [iWork runtime lane](../../OfficeIMO.IWork.Benchmarks/README.md) measures Pages, Numbers and Keynote load/projection and conversion/save workloads with saved-package semantic validation.

## AI request packing

Build `OfficeIMO.AI` in Release and run from a fresh PowerShell process on .NET 10 or newer with PSPublishModule installed:

```powershell
dotnet build OfficeIMO.AI/OfficeIMO.AI.csproj -c Release
./Build/Benchmarks/Run-AiPackingBenchmarks.ps1
```

This opt-in suite measures packing and local validation over 100, 1000 and 4000 captured short text blocks with 48,000- and 2,000,000-character request budgets. A no-inference executor isolates engine work. Validation compares every request's evidence with the original snapshot, checks complete coverage and enforces the transport budget. PowerForge runs two warmups and seven measurements per case and records managed allocations, sizing calls and request counts. `-Blocks`, `-RequestCharacters`, `-BinaryRoot`, `-OutputRoot` and the measurement-count parameters select other runs. These figures exclude source capture, inference and native OCR memory; compare the same inputs and runtime with other builds idle.
