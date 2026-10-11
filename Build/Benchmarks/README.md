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

## PDF runtime operations

Build the opt-in PDF workload and run the shared PowerForge suite from a fresh
PowerShell 7 process with PSPublishModule 3.0.168 or newer and Poppler's `pdftotext`
available:

```powershell
dotnet build OfficeIMO.Pdf.Benchmarks/OfficeIMO.Pdf.Benchmarks.csproj -c Release -f net10.0
./Build/Benchmarks/Run-PdfRuntimeBenchmarks.ps1 -OutputRoot ./Ignore/Benchmarks/PdfRuntime
```

The default 1-, 20- and 100-page matrix measures buffered/forward generation, four
inspections with fresh or reused loading, SVG/PNG rendering, and one region
replacement or movement on the middle page. Generation includes composition,
layout, serialization and the final byte-array copy. Buffered output uses PDF 1.4;
forward output uses PDF 1.7. Both retain complete layout before serialization.
Rendering covers at most 12 SVG pages or four 384×512 PNG previews. Editing includes
source loading and serialization, with the target region prepared outside timing.

Generation, inspection and rendering use the `BusinessEmbeddedCarlito` profile:
embedded regular Carlito text, a numeric table, a colored rectangle and a repeated
32×32 image. The existing benchmark font asset is loaded during preparation;
font registration remains part of generation. Editing uses the separately
labelled `TextEditStandard14` profile with four ordinary Helvetica paragraphs per
page. Mixed business-page graphics can make source text unsafe to restamp, so
that corpus does not qualify this ordinary region-edit lane. Each case records
both source fingerprints and byte counts; per-sample evidence identifies the
actual profile and input hash. Affected text spans include source space spans.

PowerForge owns timing, managed allocations, rotated ordering, and raw JSON/CSV
samples. Every measured result is validated outside timing. Poppler independently
checks all saved PDF pages, replacement absence, and retained neighboring word
geometry. Rendering requires zero capability diagnostics. PNG checks decode the
image and verify dimensions and visible content. SVG checks validate logical
drawing text and require byte-exact agreement with the public page drawing export,
including embedded cmap aliases that preserve shaped glyphs and vector rectangles
that preserve non-interpolated image pixels. These rendering
checks establish managed representation consistency, without independent visual
fidelity proof.
Poppler is a validation dependency of this opt-in suite, not a product requirement.

Use `-Pages`, `-Operation`, `-BinaryRoot`, `-WarmupCount`, `-IterationCount` and
`-Plan` to select a run. On Windows, pass `-ProcessorAffinityMask` and
`-ProcessPriority Normal` and repeat the full matrix separately for each discovered
CPU domain. The runner records requested/applied/original placement and restores
the process afterward. `-EnvironmentEvidence` accepts a JSON host snapshot for
physical/cache topology and power-plan details the runner does not discover.
Source, input, assembly, spec, reader and runner fingerprints accompany the results.
Keep other builds idle and compare only equivalent workloads and host settings.

Use `-MemorySamplingIntervalMilliseconds 5` in a separate run for sampled managed
heap and resident memory. Sampled maxima are lower bounds and include observer
effects. The output root retains exact per-iteration PDFs, rendered pages and
validation evidence; keep the final reports and decisive artifacts, then remove
superseded output. This suite adds no ordinary correctness timing thresholds.

## PDF structured reading and file destinations

The separate opt-in evidence suite selects one family explicitly and leaves the
eight-operation runtime matrix unchanged:

```powershell
./Build/Benchmarks/Run-PdfRuntimeEvidenceBenchmarks.ps1 -Family StructuredRead
./Build/Benchmarks/Run-PdfRuntimeEvidenceBenchmarks.ps1 -Family Destination
```

Use the same Release build, PowerShell 7, PSPublishModule and Poppler prerequisites
as the runtime suite. `StructuredRead` defaults to 20, 100 and 500 pages. Each
operation loads the immutable bytes afresh and performs the public structured
read with canonical parser and semantic work limits. Every page has a repeated
header whose x position and font size drift slightly, plus one ordered body
marker. Validation requires exactly one header and the complete ordered body
sequence, with independent source text/page extraction through Poppler. Source
geometry, exact limits, input fingerprints and semantic digests accompany results.

`Destination` defaults to 1, 20 and 100 pages. `DestinationDirect` composes the
business corpus and calls `Save(FileStream)`; `DestinationCopied` composes the same
corpus, calls `Save(MemoryStream)`, then `ToArray` and `File.WriteAllBytes`. Both use
the same PDF 1.7, compression and forward object serialization settings, and both
close a complete file inside measurement. They measure destination cost separately
from the buffered/forward serialization comparison. They retain complete layout
and use ordinary cached file writes, without a forced durable flush. Every saved
file is checked after timing against the exact prepared forward-output digest,
managed content/writer contracts and independent Poppler extraction on every page.
Writer store retention, budgets and spill flags are recorded separately from final
byte-array materialization.

The wrapper accepts page/operation selection, output paths, placement and host
evidence settings, and `-Plan`. PowerForge retains raw samples and outliers and
owns all timing and allocation measurement. Use sampled memory in a separate
cohort. These calls pass `CancellationToken.None` through the public APIs;
cancellation latency and storage durability are outside this measurement contract.

## PDF rendering stages and first-write observations

The phase suite uses the same Release build, shared PowerForge runner and Poppler
prerequisites. Each family is explicitly selected:

```powershell
./Build/Benchmarks/Run-PdfPhaseBenchmarks.ps1 -Family RenderStages
./Build/Benchmarks/Run-PdfPhaseBenchmarks.ps1 -Family FirstWrite -Pages 20,100,500
./Build/Benchmarks/Run-PdfPhaseBenchmarks.ps1 -Family ClosedLedger -ExpectedGeometry ClosedRuling
```

`RenderStages` uses one 20-page embedded-Carlito business source and previews pages
1–4 at scale 0.64, producing 384×512 pixels per page. `SourceOpen`, `Projection`,
`Rasterization` and `PngEncoding` measure public boundaries with their preceding
inputs prepared outside measurement; `RenderPng` measures the complete public
call. Every stage completes its remaining pipeline outside measurement and must
match the control's complete PNG bytes and decoded pixels, with zero diagnostics
and complete source text. Prepared-stage medians are not additive, and process
font/JIT/codec caches remain warm. These checks establish representation
consistency for the exact corpus; independent visual fidelity is a separate proof.

`FirstWrite` is an instrumented destination cohort. Direct and copied strategies
use identical forward PDF options and the same delegating stream around the actual
destination FileStream. Event metrics record method-entry and Save-start intervals
to the first successful nonempty write call's start and completion, its payload
length and total accepted bytes. The copied strategy saves to memory and calls
`ToArray` before writing; its first destination write contains the complete
artifact. Completion is an upper-bound observable proxy for acceptance by the
stream and OS cache, without proving the first physical byte, durable persistence
or network delivery. PowerForge alone measures total operation duration. Every
file must match the prepared digest and pass complete Poppler content readback.
Named noncollecting managed-heap/working-set boundaries and optional separate
`-MemorySamplingIntervalMilliseconds 5` observations may miss peaks and include
observer effects; they do not establish retained heap or constant-space layout.

`ClosedLedger` loads either exact independent ReportLab ledger fixture afresh and
runs public structured reading. `-Rotation 0,90` selects page rotations. Geometry
expectations are explicit: `DetectedProportions` qualifies the recorded baseline
text-derived bounds; `ClosedRuling` requires authored grid bounds and widths.
Both require the complete ordered 4×4 cell content and exact producer/input
fingerprints. Comparing these modes measures the cost of stronger recovery;
output geometry equality is deliberately outside that comparison.

Use `-Plan`, `-Operation`, paths, placement and host evidence options as in the
runtime suite. Keep these diagnostics separate from ordinary runtime cohorts and
retain losses and raw samples. Any sampled stack profile of the render control
must identify its actual tool, process and sampling semantics; font work spans
projection and glyph rasterization rather than a separate public font phase.

## AI request packing

Build `OfficeIMO.AI` in Release and run from a fresh PowerShell process on .NET 10 or newer with PSPublishModule installed:

```powershell
dotnet build OfficeIMO.AI/OfficeIMO.AI.csproj -c Release
./Build/Benchmarks/Run-AiPackingBenchmarks.ps1
```

This opt-in suite measures packing and local validation over 100, 1000 and 4000 captured short text blocks with 48,000- and 2,000,000-character request budgets. A no-inference executor isolates engine work. Validation compares every request's evidence with the original snapshot, checks complete coverage and enforces the transport budget. PowerForge runs two warmups and seven measurements per case and records managed allocations, sizing calls and request counts. `-Blocks`, `-RequestCharacters`, `-BinaryRoot`, `-OutputRoot` and the measurement-count parameters select other runs. These figures exclude source capture, inference and native OCR memory; compare the same inputs and runtime with other builds idle.
