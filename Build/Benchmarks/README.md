# Assistant evidence measurements

Build `OfficeIMO.Studio` in Release, then run the evidence suite from a fresh PowerShell 7 process:

```powershell
./Build/Benchmarks/Run-AssistantEvidenceBenchmarks.ps1
```

The PowerForge suite compares reading a PDF again with inspecting an already prepared immutable snapshot. Both paths validate the same source hash, evidence hash, page count and text coverage. The default matrix uses synthetic 1-, 100- and 500-page PDFs, two warmups and seven measured iterations per case. `-Pages`, `-WarmupCount`, `-IterationCount` and `-OutputRoot` select a different run; `-Plan` prints the execution plan.

Run from an unchanged checkout. PowerForge records source provenance and rejects runs if source changes during measurement. Keep other builds and benchmarks idle. On machines with multiple CPU or cache domains, use an explicitly recorded affinity and repeat on each relevant group.

Outputs include samples, summaries, comparison tables and environment metadata under the selected output folder. `OperationAllocatedBytes` measures managed allocations across the operation; it is not peak or retained memory. Elapsed time includes harness overhead, which matters for very short reuse operations. These measurements cover local evidence preparation and readiness only. They do not measure inference latency, model accuracy, OCR, or end-to-end UI responsiveness.
