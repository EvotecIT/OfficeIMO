# PDF runtime operations (2026-10-11)

Region replacement and movement reuse source readers within one operation. In the full-matrix confirmation, mean managed allocation fell by 1.4–2.4% in all 12 edit cases. Eleven edit medians improved; one median and two observed tails became slower. The shared host also produced large timing changes in unchanged operations, so these results support a bounded allocation reduction and observed latency differences, without establishing a causal or universal speedup.

The [companion JSON](pdf-runtime-2026-10-11.json) retains the original baseline, the initial edit comparison, the complete confirmation, structured-read and destination cohorts, all measured latency/allocation samples, input fingerprints, binary identities and raw-report hashes. It separates the two processor domains and preserves every measured loss. These are local source measurements; package publication and deployment are separate evidence.

## Work and measurement boundaries

Each input has explicit 600 × 800 point pages. Generation, inspection and rendering use `BusinessEmbeddedCarlito`: embedded Carlito text, numeric tables, a colored rectangle and a repeated 32 × 32 image. Editing uses `TextEditStandard14`: four ordinary Helvetica paragraphs per page. Mixed business graphics can make text unsafe to restamp, so the edit lane qualifies only this ordinary text profile.

- `GenerateBuffered` uses PDF 1.4; `GenerateForward` uses PDF 1.7. Both include composition, font registration, layout, compressed serialization into `MemoryStream`, and `ToArray`. The font file is prepared before measurement. Both retain complete layout; forward object serialization does not establish complete pipeline streaming.
- `InspectCold` loads four documents and inspects each once. `InspectReused` loads one document and inspects it four times within the operation. Neither keeps a document between measured iterations.
- `RenderSvg` renders the first `min(pages, 12)` pages. `RenderPng` renders the first `min(pages, 4)` pages at 384 × 512 pixels, with a 512 × 512 pixel admission ceiling. SVG permits 16,000,000 pixels per page. Rendering includes source loading and the complete public render call.
- `ReplaceRegion` and `MoveRegion` load the source, edit one prepared region on the middle page, and serialize the result. Movement is eight points horizontally. Region discovery is outside measurement; the three affected spans include the source space.
- Generation uses default object/page-content store budgets and no explicit generated-page/output-byte ceiling. Parsing uses the canonical defaults. The wrapper admits 1–500 source pages. The standard cohort has no memory sampler; allocated bytes are cumulative managed allocation, not retained heap, resident memory or a process peak.

PowerForge owns timing, process-wide managed allocation, case rotation and raw reports. Each case has two warmups and seven measured samples, with `OutlierMode=None`, memory cleanup before each iteration, and no cooldown. No observations are removed or winsorized. The reported p95 interpolates at `(n − 1) × 0.95`; seven samples do not establish a production tail.

All 672 confirmation samples passed readback outside measurement. Poppler checks every saved PDF page, complete text/numeric multiplicity, replacement absence and neighboring word geometry within 0.05 points. Movement checks both target words. Rendering requires zero capability diagnostics; PNG checks decoding, dimensions and visible content. SVG checks logical drawing text and byte-exact agreement with public drawing export, including embedded cmap aliases and nearest-neighbor image rectangles. The rendering checks establish managed representation consistency; they do not establish independent visual fidelity.

## Recorded host and source identity

The host is Windows 11 Pro build 26300, AMD Ryzen 9 9950X3D2, 16 cores/32 logical processors, High performance power plan, Normal process priority, PowerShell 7.6.6 and .NET SDK 10.0.112 with .NET 10.0.12 installed. Domain A uses affinity `0xFFFF`; B uses `0xFFFF0000`. Each discovered L3 domain has 96 MiB. Other builds and Java analysis were active. Fresh process counts before the four confirmation runs were 70, 106, 84 and 73 `dotnet` processes; counts are neither CPU utilization nor continuous contention telemetry.

The baseline assemblies identify commit `4d99473e4f4b1a57a344e129ca51621de9bb206d`. Candidate assemblies identify `b5d427f7dd9619e6972c245545ca6938d1bf92c8`, containing source-reader reuse and Save cancellation; the benchmark assembly also contains the optional evidence extension compiled before commit `f4490e8571126dc73bf580632e140db6233816d4`. The clean, frozen confirmation harness is at that latter commit. This comparison does not isolate the editor change alone. The extension leaves the eight measured operations unchanged.

The confirmation ran A baseline → A candidate → B candidate → B baseline. Each fresh process rotated the same 24 operation/page cases. Inputs, persisted iteration order, validation evidence and output hashes agreed within each domain. An independent read-only reconciliation checked 322 persisted output hashes per run and all 37 candidate snapshot files. The companion records source-file and assembly hashes separately: the harness checkout is not an archived binary's compilation identity. Original and frozen wrapper/spec files have the same Git blobs, with different file-byte hashes from checkout line endings.

Poppler is `pdftotext` 25.07.0, SHA-256 `E66E7ADA660A910918DD2974CE6486170A20C2AC73A78B003E9888B97152A82A`. Carlito is the existing SIL Open Font License 1.1 asset, SHA-256 `F6418F708BAEDE9789DAEF5D458C0F53D2A888AF9820E8062934E504FEDC6595`. PSPublishModule/PowerForge assembly hashes, full binary identities, all three input sizes/hashes and run/report fingerprints are in the companion.

## Confirmation costs at 100 source pages

These are candidate medians from the frozen confirmation, rounded for display. Allocation is the mean per complete operation in MiB (`1,048,576` bytes). SVG and PNG cover 12 and four pages respectively.

| Operation | A median, ms | B median, ms | A allocation, MiB | B allocation, MiB |
| --- | ---: | ---: | ---: | ---: |
| GenerateBuffered | 38.51 | 22.56 | 22.47 | 22.48 |
| GenerateForward | 38.10 | 22.03 | 22.36 | 22.34 |
| InspectCold, four loads | 4.45 | 2.64 | 5.92 | 5.91 |
| InspectReused, one load/four calls | 1.58 | 1.38 | 3.40 | 3.39 |
| RenderSvg, 12 pages | 87.00 | 82.78 | 91.62 | 91.79 |
| RenderPng, four previews | 43.22 | 40.45 | 37.03 | 37.03 |
| ReplaceRegion | 116.67 | 93.66 | 126.81 | 126.74 |
| MoveRegion | 110.08 | 79.52 | 126.83 | 126.76 |

The 100-page business input is 150,879 bytes; the edit input is 43,708 bytes. Single-region edits still allocate about 127 MiB because the operation loads and rewrites the document. These results identify an allocation cost for this workload, without establishing maximum supported document size or concurrent service capacity.

## Edit comparison, including losses

Negative changes mean less time or allocation. The full 24-case cadence is identical before and after in this confirmation.

| Domain | Operation/pages | Median before → after, ms | Median change | p95 change | Mean allocation change |
| --- | --- | ---: | ---: | ---: | ---: |
| A | Replace / 1 | 5.62 → 4.50 | −20.0% | −48.5% | −1.6% |
| A | Move / 1 | 5.52 → 4.15 | −24.9% | −42.6% | −1.5% |
| A | Replace / 20 | 48.84 → 35.18 | −28.0% | **+21.1%** | −1.4% |
| A | Move / 20 | 60.59 → 32.24 | −46.8% | −40.6% | −1.4% |
| A | Replace / 100 | 242.52 → 116.67 | −51.9% | −33.4% | −1.5% |
| A | Move / 100 | 216.84 → 110.08 | −49.2% | −59.9% | −1.5% |
| B | Replace / 1 | 3.17 → 3.20 | **+1.1%** | **+75.8%** | −2.4% |
| B | Move / 1 | 3.96 → 3.12 | −21.3% | −37.3% | −2.4% |
| B | Replace / 20 | 21.85 → 17.28 | −20.9% | −16.8% | −1.8% |
| B | Move / 20 | 23.07 → 18.27 | −20.8% | −4.3% | −1.8% |
| B | Replace / 100 | 112.94 → 93.66 | −17.1% | −25.2% | −1.7% |
| B | Move / 100 | 97.89 → 79.52 | −18.8% | −21.2% | −1.6% |

Domain B one-page replacement p95 rose from 5.983 to 10.518 ms. Domain A 20-page replacement p95 rose despite its lower median. Unchanged controls also shifted: 100-page forward generation was 26.8% slower and SVG rendering 34.4% slower in B, while A control medians fell by roughly 9–59% across the matrix. Host load, order, compilation/runtime differences and cache effects remain plausible contributors to the timing changes.

The initial narrower edit run remains in the companion. It rotated six candidate cases against the original 24-case baseline, had nine of 12 median losses and nine p95 losses, and showed about 1.7–10% lower mean allocation depending on size. Its candidate checkout was dirty, and its process list was reused rather than freshly sampled. Matching inputs and outputs do not remove those cadence/provenance limits. The bounded full-matrix confirmation improves comparability while preserving the original observations.

## Structured reading with drifting headers

The separate `StructuredRead` family loads a fresh PDF and calls public structured reading. Each page has a repeated header at `x = 36 + .02 × page`, font size `10 + .001 × page`, and one ordered body marker. It uses canonical parser limits, a 1,000-page semantic admission ceiling and 10,000,000 page/document work units. The companion retains the complete parser-limit snapshot. These are admitted budgets; 500 pages is the measured scale.

The candidate is frozen source `f4490e8571126dc73bf580632e140db6233816d4` plus the captured 28-line geometry-envelope change, patch SHA-256 `7942CF140539FB5ABEB2DA39C59532DA999434CABC26583EB254BF745300D0C3`. Its dirty source and seven untracked test cases are explicitly recorded; the benchmark does not compile those tests. The patch handles mutually compatible repeated-edge geometry once, preserving the pairwise fallback. Baseline assemblies have the earlier `b5d427f7` identity; candidate assemblies have the `f4490e8571` identity plus the patch. Exact binary hashes are retained.

All 84 measurements passed. Before ran A then B; candidate ran B then A, with three rotated cases and seven measured samples per case. Exact inputs, order and classification digests agreed. Poppler independently checks source text, pages and geometry; exact logical header/body assertions use the managed result, without an independent classification oracle.

| Pages | A median before → candidate, ms | A mean allocation change | B median before → candidate, ms | B mean allocation change |
| --- | ---: | ---: | ---: | ---: |
| 20 | 16.85 → 11.42 | **+2.28%** | 7.46 → 10.26 (**+37.6%**) | **+2.88%** |
| 100 | 77.57 → 50.97 | −0.88% | 35.41 → 49.02 (**+38.4%**) | −0.28% |
| 500 | 428.88 → 248.81 | −8.05% | 150.42 → 248.58 (**+65.3%**) | −7.43% |

Domain A medians fell 32–42%; every B median became slower. B p95 also rose 9.7% at 20 pages and 4.1% at 100 pages. This evidence supports lower observed allocation at 500 pages while retaining smaller-input allocation losses and substantial latency losses. Separately, the [caller-budget regressions](../../OfficeIMO.Pdf.Tests/Pdf/PdfRepeatedPageEdgeEnvelopeTests.cs) changed from two failures/five passes to seven passes after integration at `289c5e0dd8`: drifting 100-page headers/footers fit a 4,000-unit budget. That deterministic work-budget proof does not establish a timing speedup.

## Equivalent file destinations and sampled memory

`DestinationDirect` composes the business corpus and saves to `FileStream`. `DestinationCopied` composes the same corpus, saves to `MemoryStream`, calls `ToArray`, then `File.WriteAllBytes`. Both use PDF 1.7, compression, forward object serialization and complete file closure inside measurement. They use the unchanged frozen baseline snapshot; this compares existing strategies, without a product before/after change. All 84 standard samples passed exact reference-PDF digests, complete page content, independent Poppler extraction and matching writer-report checks.

| Pages | A copied → direct median, ms | A mean allocation change | B copied → direct median, ms | B mean allocation change |
| --- | ---: | ---: | ---: | ---: |
| 1 | 4.76 → 4.33 | −3.38% | 5.77 → 5.30 | −3.37% |
| 20 | 8.65 → 8.31 | −2.36% | 11.32 → 15.10 (**+33.4%**) | −2.45% |
| 100 | 22.50 → 24.54 (**+9.1%**) | −1.98% | 52.78 → 43.58 | −1.93% |

Direct output avoids the final byte-array materialization and uses 1.9–3.4% less mean managed allocation in these six standard cases. Its median latency is worse in two cases. Observed p95 also rises 2.4% and 14.1% at A's one/20 pages and 9.7% at B's 20 pages. Both strategies retain complete layout and use OS-cached writes, with no forced durable flush. At 500 pages, both report 1,843,564 retained page-content bytes, zero retained completed-object bytes and no payload-store spills; those store counters do not describe total process memory.

A separate five-millisecond memory cohort ran three measurements per strategy/domain at 100 and 500 pages; all 24 passed and no sampler failures were recorded. The table shows medians of sampled heap growth above the per-operation baseline and sampled absolute working-set maxima, in MiB. Only 4–10 observations were captured per operation. Values are lower bounds with observer effects; process activity and missed peaks limit comparisons.

| Domain/pages | Copied → direct sampled heap growth, MiB | Copied → direct sampled working-set maximum, MiB |
| --- | ---: | ---: |
| A / 100 | 22.36 → 21.88 | 255.19 → 258.99 |
| A / 500 | 59.37 → 57.62 | 256.34 → 256.35 |
| B / 100 | 22.35 → 21.90 | 260.88 → 260.09 |
| B / 500 | 59.37 → 57.62 | 257.91 → 254.33 |

Lower sampled heap growth does not imply uniformly lower resident memory: A's sampled working set is higher for direct output. The companion preserves per-iteration sampler counts, baselines, sampled maxima and every instrumented latency loss. Instrumented timings belong to this separate cohort. First-byte latency, cancellation latency, durable storage and constant-space layout are outside these measurement contracts.

## Reproduce

Use a clean checkout and the [PDF runtime instructions](../../Build/Benchmarks/README.md#pdf-runtime-operations). Build each historical source candidate separately, archive its output, and pass its archive through `-BinaryRoot` from one frozen harness. Record source and binary identities separately. Keep input profiles, tool/font hashes, placement, run order and iteration settings equal.

```powershell
dotnet build OfficeIMO.Pdf.Benchmarks/OfficeIMO.Pdf.Benchmarks.csproj -c Release -f net10.0
./Build/Benchmarks/Run-PdfRuntimeBenchmarks.ps1 -Plan
./Build/Benchmarks/Run-PdfRuntimeBenchmarks.ps1 -OutputRoot ./Ignore/Benchmarks/PdfRuntime/A -ProcessorAffinityMask 0xFFFF -ProcessPriority Normal
./Build/Benchmarks/Run-PdfRuntimeBenchmarks.ps1 -OutputRoot ./Ignore/Benchmarks/PdfRuntime/B -ProcessorAffinityMask 0xFFFF0000UL -ProcessPriority Normal
./Build/Benchmarks/Run-PdfRuntimeEvidenceBenchmarks.ps1 -Family StructuredRead -Plan
./Build/Benchmarks/Run-PdfRuntimeEvidenceBenchmarks.ps1 -Family Destination -Plan
./Build/Benchmarks/Run-PdfRuntimeEvidenceBenchmarks.ps1 -Family StructuredRead -OutputRoot ./Ignore/Benchmarks/PdfEvidence/A -ProcessorAffinityMask 0xFFFF
./Build/Benchmarks/Run-PdfRuntimeEvidenceBenchmarks.ps1 -Family Destination -OutputRoot ./Ignore/Benchmarks/PdfEvidence/A -ProcessorAffinityMask 0xFFFF
./Build/Benchmarks/Run-PdfRuntimeEvidenceBenchmarks.ps1 -Family Destination -Pages 100,500 -IterationCount 3 -MemorySamplingIntervalMilliseconds 5 -OutputRoot ./Ignore/Benchmarks/PdfMemory/A -ProcessorAffinityMask 0xFFFF
```

Supply `-EnvironmentEvidence` with a fresh host snapshot and discover the actual CPU topology before choosing affinity masks. The displayed masks describe this host. Retain raw reports and decisive output hashes; repeated full-document rendering, stronger independent visual fidelity, retained/peak-memory qualification and contention-free capacity measurements require their own evidence.
