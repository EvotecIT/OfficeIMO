# Assistant evidence measurements

Build `OfficeIMO.Studio` in Release, then run the evidence suite from a fresh PowerShell 7 process:

```powershell
./Build/Benchmarks/Run-AssistantEvidenceBenchmarks.ps1
```

The PowerForge suite compares reading a PDF again with inspecting an already prepared immutable snapshot. Both paths validate the same source hash, evidence hash, page count and text coverage. The default matrix uses synthetic 1-, 100- and 500-page PDFs, two warmups and seven measured iterations per case. `-Pages`, `-WarmupCount`, `-IterationCount` and `-OutputRoot` select a different run; `-Plan` prints the execution plan.

Run from an unchanged checkout. PowerForge records source provenance and rejects runs if source changes during measurement. Keep other builds and benchmarks idle. On machines with multiple CPU or cache domains, use an explicitly recorded affinity and repeat on each relevant group.

Outputs include samples, summaries, comparison tables and environment metadata under the selected output folder. `OperationAllocatedBytes` measures managed allocations across the operation; it is not peak or retained memory. Elapsed time includes harness overhead, which matters for very short reuse operations. These measurements cover local evidence preparation and readiness only. They do not measure inference latency, model accuracy, OCR, or end-to-end UI responsiveness.

The separate [iWork runtime lane](../../OfficeIMO.IWork.Benchmarks/README.md) measures Pages, Numbers and Keynote load/projection and conversion/save workloads with saved-package semantic validation.

## PDF redaction workflow

Build the opt-in workload and run it from a fresh PowerShell 7 process with PSPublishModule installed:

```powershell
dotnet build OfficeIMO.Pdf.Benchmarks/OfficeIMO.Pdf.Benchmarks.csproj -c Release -f net8.0
./Build/Benchmarks/Run-PdfRedactionRuntimeBenchmarks.ps1 -OutputRoot ./Ignore/Benchmarks/PdfRedactionRuntime
```

Use PSPublishModule 3.0.158 or newer. `-ModulePath` also accepts a module manifest or a source-built module DLL.

The default 1-, 25- and 100-page matrix measures precise regex search, review with edited labels, application and complete-stream verification including managed rendering. Input generation and final readback occur outside measurement. Each lane checks every page, absence of selected text, retained neighboring text and source immutability. `-Pages`, `-WarmupCount`, `-IterationCount`, `-BinaryRoot`, `-ModulePath` and `-Plan` select the run. Timing and managed allocation come from PowerForge; no measurement is part of ordinary PR correctness gates.

Use `-InputPath` with a saved PDF or a directory of PDFs and `-Pattern` to repeat the same input bytes across builds. The directory scan reads only its immediate `.pdf` files. Supplied files replace the generated page matrix and must produce exactly one reviewed area on each page. Each case records its source fingerprint, and run metadata records the effective search pattern as `SearchPatternJson` and the search options. Decode the JSON string to preserve meaningful whitespace; compare only matching fingerprints, criteria, options, runtime and host settings. Generated inputs use the fixed `private account [0-9]{3}` pattern. Complete-output verification and saved-output absence checks still run, while synthetic neighbor checks apply only to generated inputs.

Use `-MemorySamplingIntervalMilliseconds 10` in a separate run to observe managed heap and resident memory during the operation. Sampled maxima are lower bounds and include host/observer effects; they do not isolate native allocations. These workloads measure the engine calls used by Studio, rather than window responsiveness or another engine's redaction behavior. Keep other builds idle and interpret results for the recorded host and input shapes.

The [independent viewer checks](../PdfViewerVerification/README.md) use retained source/result pairs to verify extraction, geometry and rendered pages outside the measured lane.

## AI request packing

Build `OfficeIMO.AI` in Release and run from a fresh PowerShell process on .NET 10 or newer with PSPublishModule installed:

```powershell
dotnet build OfficeIMO.AI/OfficeIMO.AI.csproj -c Release
./Build/Benchmarks/Run-AiPackingBenchmarks.ps1
```

This opt-in suite measures packing and local validation over 100, 1000 and 4000 captured short text blocks with 48,000- and 2,000,000-character request budgets. A no-inference executor isolates engine work. Validation compares every request's evidence with the original snapshot, checks complete coverage and enforces the transport budget. PowerForge runs two warmups and seven measurements per case and records managed allocations, sizing calls and request counts. `-Blocks`, `-RequestCharacters`, `-BinaryRoot`, `-OutputRoot` and the measurement-count parameters select other runs. These figures exclude source capture, inference and native OCR memory; compare the same inputs and runtime with other builds idle.
