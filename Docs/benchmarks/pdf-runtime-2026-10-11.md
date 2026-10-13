# PDF runtime operations (2026-10-11)

The integrated source qualifies five rendering boundaries, first successful destination-write observations and authored closed-ruling geometry recovery. All 234 new non-profile samples and 100 separately profiled measured calls passed their output contracts. Stronger grid recovery increases measured allocation and latency; rendering and destination results retain substantial timing variation and memory losses.

Region replacement and movement reuse source readers within one operation. In the full-matrix confirmation, mean managed allocation fell by 1.4–2.4% in all 12 edit cases. Eleven edit medians improved; one median and two observed tails became slower. The shared host also produced large timing changes in unchanged operations, so these results support a bounded allocation reduction and observed latency differences, without establishing a causal or universal speedup.

The [companion JSON](pdf-runtime-2026-10-11.json) retains the original baseline, the initial edit comparison, the complete confirmation, structured-read, destination, rendering-stage, first-write, sampled-memory and closed-ledger cohorts, all measured latency/allocation samples, input fingerprints, binary identities and raw-report hashes. It separates the two processor domains and preserves every measured loss. These are local source measurements; package publication and deployment are separate evidence.

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

Lower sampled heap growth does not imply uniformly lower resident memory: A's sampled working set is higher for direct output. The companion preserves per-iteration sampler counts, baselines, sampled maxima and every instrumented latency loss. Instrumented timings belong to this separate cohort. These historical destination cohorts exclude first-byte latency, cancellation latency, durable storage and constant-space layout. The later instrumented cohort below separately observes first successful destination-write completion.

## Rendering boundaries on the integrated source

The current-head evidence uses clean product source `90e061551cee67b9b92443e273ffc2e81316c416` and a matched 37-file Release/net10.0 binary snapshot. Its PDF/Core source matches the preceding `184cd966` correctness candidate; the later integration changes the native DOC codec. The frozen phase harness is `2ed26d029163294b5b2113c979ea59530bd895ff`. Its 13 source/spec/wrapper files have the same canonical Git content as the integrated checkout; their physical CRLF/LF hashes remain separately recorded. Historical cohorts above keep their original compilation identities.

The final combined build passed a new eight-operation, 20-page functional smoke. Five rendering, two first-write and two closed-ledger smoke cases also passed. The first first-write attempt failed when Poppler tried to create XHTML at a 261-character Windows path; both 259-character PDF outputs existed and passed managed validation. The same frozen code and inputs passed full independent readback under a shorter task output root. Both failed samples and exact diagnostics remain excluded from performance statistics.

Rendering uses the same 20-page, 60,732-byte embedded-Carlito business source and four 384 × 512 previews. `SourceOpen` measures public source opening; `Projection` starts from a freshly opened reader and produces four drawings; `Rasterization` starts from newly projected drawings; `PngEncoding` starts from newly rendered RGBA rasters. Preparation and downstream artifact completion occur outside the measured operation. `RenderPng` includes source loading and the complete public render call. Inputs are fresh per sample, while process-wide JIT, font and codec caches remain warmable.

All 70 measured outputs passed zero-diagnostic, scene-geometry and complete PNG-byte/RGBA agreement checks against a fresh public `RenderPng` control. Poppler independently checked all 20 source pages and their text/geometry. PNG/control equality establishes managed representation consistency; it does not establish independent visual fidelity.

| Operation | A median / p95, ms | B median / p95, ms | A / B mean allocation, MiB |
| --- | ---: | ---: | ---: |
| SourceOpen | 0.59 / 1.70 | 0.90 / 1.69 | 2.65 / 2.63 |
| Projection, four pages | 15.19 / 66.04 | 34.54 / 41.09 | 18.74 / 18.71 |
| Rasterization, four pages | 15.76 / 189.79 | 31.20 / 151.53 | 11.49 / 11.49 |
| PngEncoding, four rasters | 15.28 / 25.95 | 19.94 / 36.15 | 2.91 / 2.90 |
| Complete RenderPng, four previews | 50.14 / 226.42 | 104.03 / 365.02 | 36.50 / 36.49 |

These separate prepared-stage medians are not additive. They do not partition one end-to-end execution or support residual-overhead subtraction. Seven samples per cell showed population coefficients of variation from 27% to 123%; the large tails remain in the companion. Fresh pre-run counts ranged from 41 to 98 `dotnet` processes, with Java and compiler processes present. The coordinator held this goal's other heavy work, but those counts do not establish CPU utilization or a contention-free host. The native cache query refreshed both 96 MiB L3 affinity domains immediately before this window.

## Sampled managed-stack attribution

A single separate domain-A profile requested ten warmups and 100 measured public `RenderPng` calls through the same PowerForge runner. All 100 measured outputs passed artifact validation; ten completed warmups are inferred from successful runner completion and have no individual readback/export. The captured EventPipe session lasted 36.008 seconds within a 45-second bound. Collection and measured work completed before the cap. The child exit code was not independently captured; qualified artifacts establish the measured scope. This profiled cohort is excluded from standard timing comparisons.

The Windows `dotnet-sampled-thread-time` capture contains 215,091 typed sample events from one PowerShell process. Of these, 3,180 have `PdfRenderStageWorkload.Execute` on their stack: 3,030 are classified Managed and 150 External. Another 211,911 samples lack that root; zero decoded samples lack a stack. Filtering excludes setup and untimed controls where the relevant frames resolve. It does not assign samples to individual iterations or establish complete native/worker-thread coverage.

| Method family present in a matched stack | Matched sample count |
| --- | ---: |
| PDF font-resource creation | 874 |
| Drawing font-family creation | 129 |
| Drawing TrueType-font methods | 360 |
| PDF TrueType Unicode cmap | 223 |
| PDF ToUnicode cmap | 277 |
| Raster RenderCore | 1,092 |
| Contour filling | 735 |
| PNG writer | 552 |
| Deflater | 355 |

Counts are inclusive unions within each explicitly named method family and overlap across rows. They are not exclusive phase shares or CPU percentages. Font preparation/cmap handling, raster contour filling and PNG compression appear in the controlled render stacks; those are concrete targets for a subsequent bounded profile/optimization. The corpus produces one named Arabic logical-text-shaper sample and no named HarfBuzz sample; absence of a named frame is not evidence of zero shaping cost.

Exactly 500 matched samples (15.7%) contain an unresolved frame: 323 have a resolved deflater caller and 177 an inflater caller. The decoder reports zero lost events and `Truncated=false`; these are parser observations, not proof of perfect collection or symbol coverage. The trace records a 1 ms sample interval, so raw counts are retained without conversion from a nominal sampling frequency into CPU time. All 2,071 distinct matched stacks, complete frames, leaf/inclusive counts, method-resolution flags and scope/tool/source hashes survive in an entry-verified lossless archive. The separate 23,722,030-byte raw trace remains retained.

## First successful destination-write observations

This separate instrumented cohort composes and closes complete direct/copied PDF 1.7 forward outputs at 20, 100 and 500 pages. It measures the first successful nonempty underlying `FileStream.Write` completion. That event is an upper-bound observable proxy for stream/OS-cache acceptance; it does not establish first physical-byte arrival, durable persistence or network delivery. PowerForge supplies total elapsed time and allocation. Benchmark-local event timestamps retain method entry, Save start, first write call start and completion, plus the call interval and payload.

All 84 standard samples passed exact reference-output digests, byte/write-count checks and independent Poppler readback of every persisted page. Direct output first accepts a 15-byte header and makes 112/432/2,032 nonempty writes. Copied output first accepts the whole 60,732/150,879/604,948-byte artifact in one write. Those payloads are different delivery milestones, even though both strategies finish with identical complete files.

| Domain / pages | Strategy | Total median / p95, ms | Method entry → first write completion, ms | Save start → first write completion, ms | First payload, bytes |
| --- | --- | ---: | ---: | ---: | ---: |
| A / 20 | Direct | 122.41 / 193.16 | 57.12 | 4.06 | 15 |
| A / 20 | Copied | 159.90 / 206.51 | 114.85 | 51.33 | 60,732 |
| A / 100 | Direct | 160.47 / 223.57 | 78.12 | 15.03 | 15 |
| A / 100 | Copied | 131.50 / 403.43 | 101.88 | 63.65 | 150,879 |
| A / 500 | Direct | 325.63 / 435.96 | 160.58 | 124.80 | 15 |
| A / 500 | Copied | 421.48 / 531.75 | 381.90 | 211.89 | 604,948 |
| B / 20 | Direct | 80.77 / 101.51 | 42.63 | 3.08 | 15 |
| B / 20 | Copied | 101.30 / 147.98 | 75.04 | 29.34 | 60,732 |
| B / 100 | Direct | 113.34 / 191.97 | 58.58 | 10.78 | 15 |
| B / 100 | Copied | 131.15 / 154.74 | 99.77 | 64.64 | 150,879 |
| B / 500 | Direct | 189.96 / 290.32 | 115.09 | 73.46 | 15 |
| B / 500 | Copied | 226.66 / 265.41 | 186.67 | 142.41 | 604,948 |

Direct Save-start to first-write completion grows from 3.08–4.06 ms at 20 pages to 73.46–124.80 ms at 500 pages. A small first header does not establish a constant-time first-write boundary or incremental layout; both strategies still compose and save the complete document.

All event medians are from the unsampled instrumented cohort. Named memory observations call `GC.GetTotalMemory(false)` and query the current process working set inside the measured operation. Direct first-write completion is captured before its memory callback. Copied output additionally observes `SaveComplete` and `BeforeDestinationWrite` before its first destination write. These observations add overhead and differ by strategy; the table does not estimate uninstrumented production latency, isolated composition cost or a causal cost of copying. The first call-start/completion interval and every raw timestamp remain in the companion.

Direct output used 1.74–2.19% less mean managed allocation in the six standard comparisons. Its A/100-page total median rose 22.0%; B/100- and B/500-page p95 rose 24.1% and 9.4%. The separate sampled-memory cohort also retained an A/100-page p95 increase of 0.09%, and B/100-page median/p95 increases of 0.33%/0.58%. These observations are comparisons of the instrumented strategies, without a product before/after change.

## Named and periodically sampled memory

The event cohort retains six named snapshots for direct output and seven for copied output, including copied `BeforeDestinationWrite` after `ToArray`. Their actual order, absolute heap/working-set values and signed deltas from the same iteration's entry snapshot remain in the JSON. Boundaries have no phase timestamps, and heap differences are not phase allocation. At 500 pages, median heap growth at `CompositionComplete` was about 7.26–7.27 MiB for both strategies/domains; at `DestinationClosed` it was 28.98–28.99 MiB for direct and 30.71–30.73 MiB for copied. Collection and host activity can reduce later observations.

A separate 5 ms periodic-memory cohort ran two warmups and three measured samples per strategy/domain at 100 and 500 pages. All 24 passed, with 6–11 observations per operation and zero recorded sampler failures. These are medians of sampled heap growth above each operation's baseline and sampled absolute working-set maxima, in MiB.

| Domain / pages | Copied → direct sampled heap growth, MiB | Copied → direct sampled working-set maximum, MiB |
| --- | ---: | ---: |
| A / 100 | 12.73 → 23.28 | 240.84 → 264.50 |
| A / 500 | 54.33 → 43.01 | 266.64 → 255.38 |
| B / 100 | 23.75 → 23.30 | 261.05 → 268.44 |
| B / 500 | 48.88 → 55.48 | 272.70 → 271.58 |

Direct output has higher observed heap growth in A/100 and B/500, and higher working-set maxima in both 100-page cases. The companion preserves all 24 sampler counts, errors, baselines, maxima/deltas and named snapshots. These observations can miss peaks and include observer effects; they establish neither retained heap, true peak memory, native allocation attribution nor constant-space layout. Their latency/allocation distributions remain separate from the preceding unsampled event cohort and the historical ordinary destination cohort.

## Cost of authored closed-ruling geometry recovery

The closed-ledger workload freshly loads each immutable independent-producer fixture at page rotations 0 and 90, then calls public structured reading. Baseline product source is `e8c8321b3dc3baca01d84aaab38333266fcc7e66` with the same benchmark-only harness. Its actual assemblies have plain informational versions; source/build and all 37 binary hashes establish its identity separately. The integrated after source is `90e061551cee67b9b92443e273ffc2e81316c416`, including the reviewed overpaint-safe grid repair and other qualified product changes. The comparison therefore describes observed cost of stronger recovery and has no causal attribution to the grid patch alone.

All 56 measurements preserve the same ordered 4 × 4 labels and semantic digest. Baseline detects column widths 166.338/164/100/81.662 points and a text-derived bounding grid. The repaired output recovers authored widths 240/72/96/120 points and the authored bounding grid. Geometry and its output digest intentionally differ; equivalent-output speedup is not the contract. Poppler independently qualifies source text and rotated word order, while producer metadata and exact managed geometry assertions qualify the grid.

| Domain / rotation | Median before → after, ms | p95 before → after, ms | Mean allocation before → after, KiB | Median / allocation change |
| --- | ---: | ---: | ---: | ---: |
| A / 0 | 4.56 → 6.43 | 5.23 → 7.81 | 3,206.75 → 3,470.10 | +41.1% / +8.2% |
| A / 90 | 4.44 → 4.56 | 6.71 → 6.10 | 3,204.56 → 3,470.57 | +2.6% / +8.3% |
| B / 0 | 2.87 → 3.45 | 3.44 → 3.80 | 3,176.48 → 3,477.11 | +20.2% / +9.5% |
| B / 90 | 2.84 → 4.05 | 3.62 → 5.65 | 3,174.30 → 3,471.62 | +42.8% / +9.4% |

All four medians and allocation means increase. Three p95 values increase; A/90 improves 9.0%. Runs use A before → after and B after → before, with matching fixture, parser/work limits, harness, rotation cadence and seven retained samples per cell. Broader integration, finite samples and shared-host effects limit interpretation.

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
./Build/Benchmarks/Run-PdfPhaseBenchmarks.ps1 -Family RenderStages -Plan
./Build/Benchmarks/Run-PdfPhaseBenchmarks.ps1 -Family FirstWrite -Plan
./Build/Benchmarks/Run-PdfPhaseBenchmarks.ps1 -Family ClosedLedger -ExpectedGeometry ClosedRuling -Plan
./Build/Benchmarks/Run-PdfPhaseBenchmarks.ps1 -Family RenderStages -OutputRoot (Join-Path $env:TEMP 'pdf-phase/A') -ProcessorAffinityMask 0xFFFF
./Build/Benchmarks/Run-PdfPhaseBenchmarks.ps1 -Family FirstWrite -Pages 20,100,500 -OutputRoot (Join-Path $env:TEMP 'pdf-write/A') -ProcessorAffinityMask 0xFFFF
./Build/Benchmarks/Run-PdfPhaseBenchmarks.ps1 -Family FirstWrite -Pages 100,500 -IterationCount 3 -MemorySamplingIntervalMilliseconds 5 -OutputRoot (Join-Path $env:TEMP 'pdf-memory/A') -ProcessorAffinityMask 0xFFFF
```

Supply `-EnvironmentEvidence` with a fresh host snapshot and discover the actual CPU topology before choosing affinity masks. The displayed masks describe this host. Retain raw reports and decisive output hashes; repeated full-document rendering, stronger independent visual fidelity, retained/peak-memory qualification and contention-free capacity measurements require their own evidence.

Use a short output root for native Windows Poppler readback; the failed 261-character XHTML path is retained as excluded qualification evidence. Historical binary/TRX copies and both 37-file phase binary snapshots are preserved in entry-verified archives through the companion's retention mapping. Completed runtime artifacts and raw samples also have an entry-verified archive mapping from their original paths. Current profile evidence retains the raw trace and a lossless complete-frame archive.

## Document-owned font reuse: separate renderer comparison

This comparison isolates the renderer change from baseline `b12d7aabf680d254ef0b090f9682257ef37014e0` to candidate `22534e8708e85ea8dc65e0dfeda86786d0c09a9d`, using canonical harness `2ed26d029163294b5b2113c979ea59530bd895ff`. The candidate reuses font decoding state within one parsed document, avoids unused decoder/width-provider construction on render paths, and bounds admission to the font-resource cache. These 168 measurements form a separate cohort from the earlier historical results, including the 90e061 rendering and 334 phase/profile observations. Later integrated Word and layout repairs are outside this comparison.

Both snapshots use byte-identical embedded-Carlito business PDFs at 1, 20 and 100 source pages, the same 17 canonical harness files, and matching tool/font hashes. `RenderSvg` renders the first `min(pages, 12)` pages; `RenderPng` renders the first `min(pages, 4)` pages at 384 × 512 pixels with a 512 × 512 admission ceiling. SVG retains the 16,000,000-pixel per-page ceiling. Each operation includes source loading and the complete public render call; 100 source pages therefore do not mean 100 rendered previews.

The four runs use A before → after and B after → before, two warmups and seven retained samples per operation/page/domain cell. All 168 measured records succeeded. All 952 encoded SVG/PNG artifacts match their before/after controls; all 252 PNG artifacts also match in complete 384 × 512 RGBA dimensions and bytes. Independent Poppler readback qualifies every source page. These checks establish compatibility with the baseline renderer, while independent visual fidelity remains a separate requirement.

Before → after values below retain both processor domains. Allocation is the mean cumulative managed allocation in MiB. Positive changes are measured losses.

| Domain / operation / source pages | Mean allocation before → after, MiB | Allocation change | Median before → after, ms | Median change | p95 before → after, ms | p95 change |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| A / RenderSvg / 1 | 11.18 → 9.51 | −14.99% | 10.75 → 6.28 | −41.60% | 19.45 → 26.77 | +37.65% |
| A / RenderSvg / 20 | 92.57 → 67.20 | −27.40% | 91.69 → 56.64 | −38.23% | 197.66 → 302.68 | +53.13% |
| A / RenderSvg / 100 | 92.97 → 67.87 | −27.00% | 123.36 → 75.87 | −38.50% | 257.54 → 222.01 | −13.80% |
| A / RenderPng / 1 | 12.12 → 10.46 | −13.69% | 28.32 → 31.52 | +11.31% | 33.51 → 59.34 | +77.09% |
| A / RenderPng / 20 | 37.46 → 29.36 | −21.64% | 144.22 → 130.52 | −9.50% | 160.73 → 210.52 | +30.97% |
| A / RenderPng / 100 | 37.98 → 29.89 | −21.31% | 115.45 → 143.22 | +24.06% | 192.58 → 247.84 | +28.69% |
| B / RenderSvg / 1 | 11.60 → 9.49 | −18.14% | 14.03 → 8.25 | −41.22% | 16.87 → 52.72 | +212.62% |
| B / RenderSvg / 20 | 92.81 → 67.19 | −27.61% | 141.95 → 133.62 | −5.87% | 208.97 → 226.12 | +8.20% |
| B / RenderSvg / 100 | 93.45 → 67.79 | −27.46% | 154.15 → 86.96 | −43.59% | 219.46 → 268.71 | +22.44% |
| B / RenderPng / 1 | 12.55 → 10.48 | −16.45% | 33.32 → 36.40 | +9.24% | 59.27 → 62.71 | +5.81% |
| B / RenderPng / 20 | 37.89 → 29.34 | −22.57% | 130.83 → 146.82 | +12.22% | 167.81 → 213.28 | +27.09% |
| B / RenderPng / 100 | 38.43 → 29.88 | −22.26% | 158.44 → 154.08 | −2.75% | 220.22 → 274.96 | +24.86% |

Mean managed allocation decreases in all twelve cells: 14.99–27.61% for SVG and 13.69–22.57% for PNG. All six SVG medians decrease by 5.87–43.59%. PNG has four median losses, from 9.24% to 24.06%, and two improvements. Observed p95 increases in eleven of twelve cells; B/SVG/1 rises 212.62%, while A/SVG/100 decreases 13.80%. Sample coefficients of variation increase in all twelve comparisons. The result supports a bounded allocation improvement with equivalent output; it does not establish a general rendering-latency improvement.

Other coordinated heavy validation was paused during the paired measurement windows. Fresh process snapshots still contain 48/131/131/98 `dotnet` processes in A-before/A-after/B-after/B-before order, plus three compiler processes in each case. Those counts are neither CPU utilization nor continuous contention telemetry. The runs retain the two native 96 MiB L3 domains, masks `0xFFFF` and `0xFFFF0000`, and Normal priority. Seven samples and shared-host variation do not establish production tails or uncontended capacity. All outliers remain included; p95 uses sorted linear interpolation at `(n − 1) × 0.95`.

PowerForge’s allocation counter covers process-wide managed bytes. It does not measure exact CPU time, native allocation, retained heap, true peak memory or the cache’s final retained size. Cache admission permits up to 64 font resources and an estimated 64 MiB retained-font budget per parsed document. Resources beyond admission remain usable without cache retention. This estimate is not a hard heap cap, and page/drawing objects can retain resources independently.

The companion’s separate `rendererComparison` object preserves all twelve exact summaries, the 168 raw timing/allocation observations, every measured loss, input fingerprints and the compact provenance index. Its source analysis is `renderer-comparison-analysis.json`, SHA-256 `D51724346B8B7FAC03187F055793ADF16AEC0D322CA9E990FF87A58DFCA98B07`, with 1,296 referenced evidence files and complete encoded/pixel mappings. Existing cohort statistics and source identities remain unchanged.

A separate resource-audit replay is retained in the companion’s archive mapping: 14 entry-verified files, 9,485,414 compressed bytes and 25,612,488 expanded bytes. Its before/after DLLs retain their documented informational-version limitation and are separate from the frozen performance snapshots. The cleanup receipt records 1,137,128,631 logical temporary bytes removed; it makes no physical-reclamation claim. The frozen renderer binaries, 168 measured records and their encoded/pixel evidence remain retained.

The complete renderer comparison is retained in the companion’s `rendererComparison.retention.measuredComparisonArchive` mapping: a 123,593,478-byte ZIP containing both 37-file frozen binary snapshots and the complete measured-output tree, with 1,267 files, 392 directories and 389,512,293 raw file bytes. The entry map resolves each original absolute path to its archive entry and preserves empty directories. All 1,266 analysis references within those source trees match the mapped lengths and hashes. The archive also retains the failed `a0` preflight receipt, which remains excluded from the 168 measured records. Archiving leaves the immutable analysis, hash-pinned qualification receipts and measured results unchanged. The separate cleanup receipt confirms removal of the three expanded evidence trees, totaling 389,512,293 logical bytes; it makes no physical-reclamation claim.


## Final integrated checkpoint — 2026-10-12

This checkpoint compares frozen product `90e061551cee67b9b92443e273ffc2e81316c416` with the completed integration at `43b3c8b4fe7b4c8918155b78741ab7ab73614c9f`, using the same `2ed26d029163294b5b2113c979ea59530bd895ff` harness. It is a separate broad-source cohort. The historical results above and the isolated 168-record renderer comparison retain their original identities and values.

All 924 measured observations passed output validation in 12 wrapper calls: 672 runtime, 84 structured-read and 168 destination-write records. The analysis reconciles 66 paired operation/size/domain comparisons and all raw observations. Prepared input variables and fingerprints match in every comparison. Timing covers the declared operation; input preparation, readback and output validation remain outside it. This does not attribute a broad final-source change to one patch or establish pixel equivalence for every rendered output.

Rendering mean managed allocation is lower in all 12 comparisons, by 18.56–28.03%. Across the complete matrix, median latency is lower in 23 cells and higher in 43; observed p95 is lower in 21 and higher in 45; mean allocation is lower in 25 and higher in 41. All twelve destination-write medians are higher. Five of six structured-read medians are higher, including the 100-page domain-A observation of 31.3075 to 119.1141ms (+280.47%). These losses remain part of the result. The cohort supports a bounded rendering-allocation improvement, not a universal latency improvement.

Fresh native cache evidence identifies two group0 96 MiB L3 domains: A uses `0xFFFF`; B uses `0xFFFF0000`. Normal priority, two warmups, seven measurements per cell and rotated operation order apply on both sides. A runs baseline then final; B runs final then baseline. Coordinated heavy work was stopped before the capture at 01:33:16 UTC and held until execution finished at 01:38:05 UTC. This remains a shared-host diagnostic: process counts are not CPU utilization or continuous contention telemetry, and seven samples do not establish production tails or capacity. No outlier is removed.

The final graph contains 37 files/137,139,639 bytes; baseline90 contains 37 files/137,122,452 bytes. Their manifests, all 17 harness blobs, actual run reports and saved output evidence are bound to the analysis. Independent text validation uses Poppler `pdftotext` 25.07.0 SHA `E66E7ADA660A910918DD2974CE6486170A20C2AC73A78B003E9888B97152A82A`; the separate layout image check uses the admitted `pdftoppm` 26.07.0 executable. The two tools are distinct.

The following tables retain every paired cell. Negative change means lower measured cost. Allocation is process-wide cumulative managed allocation, not retained heap, native allocation or true peak memory. SVG renders the first min(pages,12) pages; PNG renders the first min(pages,4) at 384×512. Destination completion observes stream/OS-cache acceptance, not durable storage delivery.

### CPU domain A

| Operation | Source pages | Median ms:90 → final | Median change | Observed p95 change | Mean allocation change | Final mean allocation MiB |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `GenerateBuffered` | 1 | 2.1173 → 2.4494 | +15.69% | +8.50% | +0.24% | 4.661 |
| `GenerateForward` | 1 | 2.1064 → 2.3399 | +11.09% | -48.88% | +0.23% | 4.733 |
| `InspectCold` | 1 | 0.7789 → 0.7881 | +1.18% | +69.28% | +0.37% | 2.901 |
| `InspectReused` | 1 | 0.4160 → 0.4532 | +8.94% | +5.51% | +0.45% | 2.365 |
| `RenderSvg` | 1 | 6.3204 → 10.8424 | +71.55% | +107.50% | -20.37% | 7.962 |
| `RenderPng` | 1 | 11.1421 → 12.7299 | +14.25% | +19.05% | -18.56% | 8.933 |
| `ReplaceRegion` | 1 | 3.7617 → 4.3493 | +15.62% | +100.34% | +0.17% | 5.170 |
| `MoveRegion` | 1 | 3.7057 → 4.2047 | +13.47% | +270.97% | +0.18% | 5.191 |
| `GenerateBuffered` | 20 | 5.4329 → 7.3794 | +35.83% | +16.56% | +0.09% | 8.320 |
| `GenerateForward` | 20 | 5.3166 → 16.2274 | +205.22% | +142.77% | +0.10% | 8.344 |
| `InspectCold` | 20 | 1.3808 → 1.7258 | +24.99% | +153.93% | +0.31% | 3.491 |
| `InspectReused` | 20 | 0.6658 → 0.5897 | -11.43% | -11.27% | +0.46% | 2.569 |
| `RenderSvg` | 20 | 74.0097 → 71.7760 | -3.02% | +43.51% | -27.93% | 65.602 |
| `RenderPng` | 20 | 40.5447 → 71.1271 | +75.43% | +64.06% | -23.36% | 27.790 |
| `ReplaceRegion` | 20 | 22.6440 → 26.1671 | +15.56% | +3.45% | -0.05% | 28.449 |
| `MoveRegion` | 20 | 23.3519 → 30.5744 | +30.93% | -5.72% | -0.05% | 28.469 |
| `GenerateBuffered` | 100 | 18.5728 → 27.0275 | +45.52% | +23.42% | -0.04% | 23.739 |
| `GenerateForward` | 100 | 18.6291 → 26.3268 | +41.32% | +32.22% | -0.04% | 23.622 |
| `InspectCold` | 100 | 3.7593 → 3.9571 | +5.26% | +24.28% | +0.17% | 5.833 |
| `InspectReused` | 100 | 1.4329 → 1.5236 | +6.33% | -16.56% | +0.30% | 3.317 |
| `RenderSvg` | 100 | 66.9184 → 59.8702 | -10.53% | +55.10% | -27.96% | 65.959 |
| `RenderPng` | 100 | 42.8761 → 38.9797 | -9.09% | +64.84% | -23.10% | 28.291 |
| `ReplaceRegion` | 100 | 105.7254 → 125.5134 | +18.72% | +29.06% | -0.16% | 126.429 |
| `MoveRegion` | 100 | 95.6195 → 129.7295 | +35.67% | +22.05% | -0.17% | 126.447 |
| `StructuredRead` | 20 | 5.7426 → 21.4082 | +272.80% | +217.40% | +0.27% | 6.357 |
| `StructuredRead` | 100 | 31.3075 → 119.1141 | +280.47% | +196.63% | +0.52% | 23.256 |
| `StructuredRead` | 500 | 182.0173 → 633.5833 | +248.09% | +153.78% | +0.55% | 107.329 |
| `DestinationDirect` | 1 | 4.7672 → 5.7283 | +20.16% | +11.64% | -0.15% | 4.567 |
| `DestinationCopied` | 1 | 4.6292 → 5.2119 | +12.59% | +11.94% | -0.14% | 4.732 |
| `DestinationDirect` | 20 | 11.5045 → 23.9754 | +108.40% | -2.32% | +0.73% | 8.237 |
| `DestinationCopied` | 20 | 18.2232 → 23.0048 | +26.24% | +15.93% | +0.46% | 8.427 |
| `DestinationDirect` | 100 | 83.7159 → 98.3612 | +17.49% | +8.10% | +0.97% | 23.691 |
| `DestinationCopied` | 100 | 80.3896 → 112.3326 | +39.74% | +3.21% | +0.95% | 24.129 |

### CPU domain B

| Operation | Source pages | Median ms:90 → final | Median change | Observed p95 change | Mean allocation change | Final mean allocation MiB |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `GenerateBuffered` | 1 | 2.2078 → 2.2501 | +1.92% | +47.20% | +0.01% | 4.651 |
| `GenerateForward` | 1 | 2.1938 → 2.2981 | +4.75% | +282.34% | -0.00% | 4.723 |
| `InspectCold` | 1 | 0.7735 → 0.7844 | +1.41% | -16.50% | -0.02% | 2.891 |
| `InspectReused` | 1 | 0.3841 → 0.4707 | +22.55% | +666.68% | -0.03% | 2.355 |
| `RenderSvg` | 1 | 7.3733 → 4.6249 | -37.28% | -43.21% | -20.56% | 7.947 |
| `RenderPng` | 1 | 31.3096 → 10.0019 | -68.05% | +122.00% | -18.74% | 8.917 |
| `ReplaceRegion` | 1 | 3.9030 → 3.7564 | -3.76% | +7.54% | +0.00% | 5.160 |
| `MoveRegion` | 1 | 4.0119 → 4.1051 | +2.32% | -21.42% | +0.00% | 5.181 |
| `GenerateBuffered` | 20 | 6.2542 → 5.4046 | -13.58% | +79.59% | +0.01% | 8.317 |
| `GenerateForward` | 20 | 9.1301 → 7.8881 | -13.60% | +62.10% | +0.01% | 8.339 |
| `InspectCold` | 20 | 1.9038 → 1.3579 | -28.67% | -54.52% | -0.02% | 3.481 |
| `InspectReused` | 20 | 0.9192 → 0.5910 | -35.70% | +55.00% | -0.02% | 2.559 |
| `RenderSvg` | 20 | 97.9181 → 45.6155 | -53.41% | -51.73% | -28.03% | 65.538 |
| `RenderPng` | 20 | 51.5027 → 32.8451 | -36.23% | -19.59% | -23.46% | 27.768 |
| `ReplaceRegion` | 20 | 25.1022 → 18.3599 | -26.86% | -17.95% | +0.12% | 28.464 |
| `MoveRegion` | 20 | 21.2907 → 19.4595 | -8.60% | +58.98% | +0.12% | 28.485 |
| `GenerateBuffered` | 100 | 24.1411 → 18.3827 | -23.85% | -39.98% | +0.07% | 23.764 |
| `GenerateForward` | 100 | 18.5999 → 17.9066 | -3.73% | -10.52% | +0.08% | 23.648 |
| `InspectCold` | 100 | 3.8969 → 3.8223 | -1.91% | -8.01% | +0.03% | 5.825 |
| `InspectReused` | 100 | 1.6449 → 1.2902 | -21.56% | -10.17% | +0.05% | 3.310 |
| `RenderSvg` | 100 | 74.6383 → 49.7824 | -33.30% | -35.64% | -27.90% | 66.038 |
| `RenderPng` | 100 | 45.6740 → 48.1162 | +5.35% | -24.75% | -23.12% | 28.292 |
| `ReplaceRegion` | 100 | 138.3351 → 123.7709 | -10.53% | +7.06% | +0.17% | 126.639 |
| `MoveRegion` | 100 | 127.1490 → 119.0660 | -6.36% | +29.52% | +0.17% | 126.660 |
| `StructuredRead` | 20 | 6.8980 → 10.6281 | +54.08% | -10.94% | +0.49% | 6.352 |
| `StructuredRead` | 100 | 35.4406 → 42.4490 | +19.78% | -19.13% | +0.18% | 23.200 |
| `StructuredRead` | 500 | 219.7094 → 212.7586 | -3.16% | -30.63% | +0.11% | 107.012 |
| `DestinationDirect` | 1 | 4.0248 → 5.2241 | +29.80% | +3.38% | +0.00% | 4.578 |
| `DestinationCopied` | 1 | 4.4407 → 4.6411 | +4.51% | +35.35% | +0.03% | 4.742 |
| `DestinationDirect` | 20 | 8.5136 → 20.9478 | +146.05% | +59.21% | +0.81% | 8.191 |
| `DestinationCopied` | 20 | 9.1407 → 18.0296 | +97.25% | +95.09% | +0.83% | 8.390 |
| `DestinationDirect` | 100 | 23.6638 → 72.3555 | +205.76% | +95.38% | +1.76% | 23.485 |
| `DestinationCopied` | 100 | 24.9446 → 82.0201 | +228.81% | +122.13% | +2.47% | 23.997 |

### Correctness and native scope of the final source

Nine fresh source-bound PDF/Word/Excel lanes across net10.0, net8.0 and net472 pass 44,774 tests with zero failures and 18 actual expected COM opt-in skips. All 948 selected critical executions pass exactly once; nine TRXs and 39 physical DLL records bind the proof to 43b. The explicit SourceRevisionId property is controller-supplied and supported by clean Git checks and physical DLL/TRX hashes. Earlier failed attempts and their original producers remain retained; their results are not relabelled as final-head execution.

All five final PDF layout artifacts match the qualified geometry and PNGs under the admitted renderer, and all five were visually inspected. The separate Word comparison has 50 passing executions over 49 source/native pairs: 114 complete PDF pages per side and 108 selected visual pages in 50 inspected contact sheets. Ordered text, numeric order/multiplicity/page assignment and native glyph origins pass their unchanged 0.30 pt X/0.25 pt Y limits; maxima are 0.162003 pt X/0.179993 pt Y. The three earlier body-spacing failures pass, as do both body pages in the tall 48/64 pt controls.

This is finite fixture evidence. The tall controls retain native width 260.040009 pt versus actual 260 pt at equal 300 pt height; exact page-size equality is not claimed. Watermark text proportions and image opacity still visibly differ. General font fidelity, including Cambria Math, and the separate 100 pt physical-margin overflow remain open. Generated NUMBER payload has a separately guarded caller contract without a native oracle; its admitted 0.01 pt horizontal bbox precision does not weaken native glyph, origin, page, text or multiplicity guards.

The companion's `finalCheckpoint` object retains the complete 924 raw observations, all 66 summaries/losses, source/tool/environment bindings, output-validation records and the bounded quality references. Source analysis is `final-runtime-checkpoint-analysis.json`, SHA `4C43DB1CA419394DA52358D4DD23F690504CD2BBF3C88080731C1C17891EFA35`; actual checkpoint qualification is SHA `551D835175A794103CC348FD532DA513042EAE848ACEA00DA7BC601AD5816D15`. Reproduction uses the existing runtime/evidence wrappers with the declared sizes, frozen binaries, affinity masks, two warmups and seven iterations; the earlier 334-record, profile and 168-record renderer cohorts are not rerun or pooled with this checkpoint.


The final cohort's retained archive is `final-qualified-evidence-43b3c8b4-20261012T015843Z.zip`, 164,877,976 bytes, SHA `DB6E5B3CDCD4EA8BF5BAAD0B3EAF6DA7EA518B58B8C9243139D59070B5C2E411`. Independent verification checks all 5,329 entries (3,665 files and 1,664 directories), 597,902,774 decompressed bytes, and all 3,688 mapped original files. The embedded `archive-control/final-qualified-evidence-43b-membership.jsonl` resolves original task-relative paths and 24 content-identical sparse aliases to their exact payloads. It preserves the complete final runtime/output graph, source-bound gate/native/layout producers and control evidence. Baseline90 stays in its existing archive, with all 37 mapped entries freshly verified; it is not duplicated. The companion's `finalCheckpoint.retention` binds the archive, map, execution, release and independent admission. Final PDF/Word visual deliverables and original native references also remain available in place. Sparse producer preservation is not a self-contained test build: the declared .NET runtime, native analysis/render tools and installed fonts remain external. Fresh cleanup/accounting receipts are separate from this archive identity and do not establish physical disk reclamation.


## Structured reading and saving follow-up — 2026-10-12

Guarded reuse of PDF operator observations reduces repeated parser work during structured reading. Bounded reuse of immutable font measurements reduces repeated text work during saving, while each document retains its own glyph usage, subset and Unicode mappings. Writer kerning bookkeeping is allocated only when needed. The measured allocation reduction is consistent across the original comparison and its one reproducibility repeat. Elapsed-time results vary substantially between those cohorts.

This follow-up compares the previously qualified product `43b3c8b4fe7b4c8918155b78741ab7ab73614c9f` with `e01dcd8f984a54749b72e3e67b36640853722bee`. All 17 canonical harness files remain at `2ed26d029163294b5b2113c979ea59530bd895ff`. The baseline snapshot has 37 files/137,139,639 bytes; the candidate has 37 files/137,142,233 bytes. Legacy directory labels containing `90` identify the baseline role here; its actual product source is `43b3`, not the older `90e0` baseline used in the preceding checkpoint.

The original full comparison ran at 11:10–11:14 UTC on 12 October and retains 924 observations across 66 paired cells. All output checks passed, and paired inputs were identical. Across all families, 28 median latencies were lower and 38 higher; mean managed allocation was lower in 53 cells and higher in 13. All 18 targeted structured-read and destination-save medians were higher. The companion JSON preserves every control-family cell and raw sample as `structuredSavingFollowup.original924`.

One bounded repeat of the targeted families ran at 12:35:00–12:37:18 UTC, after an explicit extension of the coordinating agent's quiet window through 12:40 UTC. It retains 252 observations, eight calls and 18 paired cells. The source, frozen binaries, input bytes, fonts, output validation and original measurement settings were unchanged: two warmups, seven measurements, Normal priority, rotated operation order, no removed outliers, and both freshly verified 96 MiB L3 domains (`0xFFFF` and `0xFFFF0000`). Domain A runs baseline then candidate; B reverses product order. The coordinating agent kept its own heavy work idle; this did not establish that the whole machine was idle.

In the repeat, 17 of 18 targeted medians were lower and one was higher. Structured-read medians were 34.10–58.02% lower; 100-page destination-save medians were 12.13–34.80% lower. Direct saving of one page on domain B remained 8.40% higher, or 4.6792 → 5.0723 ms. Domain A's one-page direct-save observed p95 was also 18.13% higher despite its lower median. Every repeat cell allocated less managed memory: structured reading decreased by 22.26–32.78%, and 100-page saving by 9.17–9.72%. The original allocation reductions were about 22–33% and 7%, respectively.

The table keeps both timing directions and every targeted pair. Negative change means lower cost; observed p95 comes from seven samples and is not a production tail guarantee.

| Domain | Operation | Pages | Original median change | Repeat median ms: before → after | Repeat median change | Repeat observed p95 change | Mean allocation change: original / repeat |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| A | `StructuredRead` | 20 | +57.67% | 25.9090 → 10.8756 | -58.02% | -76.88% | -21.69% / -22.35% |
| A | `StructuredRead` | 100 | +140.83% | 136.4740 → 78.4442 | -42.52% | -17.37% | -29.84% / -30.30% |
| A | `StructuredRead` | 500 | +90.48% | 576.8858 → 355.7425 | -38.33% | -29.14% | -32.38% / -32.78% |
| A | `DestinationCopied` | 1 | +22.56% | 4.7910 → 4.5282 | -5.49% | -6.94% | -0.21% / -0.23% |
| A | `DestinationCopied` | 20 | +81.11% | 8.8707 → 7.3882 | -16.71% | -4.15% | -3.83% / -4.86% |
| A | `DestinationCopied` | 100 | +194.58% | 33.2181 → 21.6581 | -34.80% | -23.81% | -6.68% / -9.26% |
| A | `DestinationDirect` | 1 | +0.20% | 4.9959 → 4.6292 | -7.34% | +18.13% | -0.26% / -0.05% |
| A | `DestinationDirect` | 20 | +52.17% | 11.5426 → 9.6164 | -16.69% | -37.71% | -3.87% / -5.09% |
| A | `DestinationDirect` | 100 | +163.92% | 26.0899 → 20.4673 | -21.55% | -28.48% | -7.14% / -9.17% |
| B | `StructuredRead` | 20 | +36.01% | 14.2198 → 8.0221 | -43.59% | -49.15% | -23.83% / -22.26% |
| B | `StructuredRead` | 100 | +1.51% | 63.6859 → 38.0474 | -40.26% | -40.98% | -30.76% / -30.29% |
| B | `StructuredRead` | 500 | +5.98% | 358.0069 → 235.9328 | -34.10% | -35.78% | -33.02% / -32.73% |
| B | `DestinationCopied` | 1 | +4.95% | 4.8384 → 4.7273 | -2.30% | -6.74% | -0.31% / -0.31% |
| B | `DestinationCopied` | 20 | +62.74% | 14.4497 → 10.5138 | -27.24% | -28.48% | -4.23% / -5.09% |
| B | `DestinationCopied` | 100 | +48.08% | 41.7660 → 36.6982 | -12.13% | -26.42% | -7.25% / -9.72% |
| B | `DestinationDirect` | 1 | +0.86% | 4.6792 → 5.0723 | +8.40% | +4.89% | -0.35% / -0.38% |
| B | `DestinationDirect` | 20 | +80.88% | 11.9467 → 10.1807 | -14.78% | -3.88% | -4.12% / -5.45% |
| B | `DestinationDirect` | 100 | +47.72% | 55.9527 → 41.3061 | -26.18% | -33.94% | -7.38% / -9.69% |

The allocation profiles identify repeated parser/state/string work during structured reading and repeated font measurement and glyph-processing work during saving. They attribute sampled allocation ticks over full processes, including preparation and validation. They are approximate diagnostics rather than latency measurements. In the original diagnostic cohort, observed collection counts decreased from 123 to 104 for reading but increased from 284 to 327 for saving; the saving increase remains a loss despite lower sampled allocation.

Separate diagnostics retain 126 observations with twenty warmups, 28 instrumented canonical measurements, and single longer repeated-loop pairs. More warmups reduced most original gaps but left mixed latency: 20- and 100-page reading stayed about 12% slower, 500-page reading was about 1% faster, and 100-page direct saving was about 40% slower. Instrumentation reversed the original timing direction. Long-loop process CPU changed by about −1.56% for reading and +0.74% for saving while elapsed time rose about 19.1% and 1.51%. No JIT cause was established because the captured trace level contained no JIT events. These diagnostics remain separate and are not pooled with either benchmark cohort.

The broad original latency regression did not reproduce with unchanged inputs and code. Lower allocation is supported consistently; a uniform latency, CPU, throughput or capacity improvement is not established. The original losses, the repeat's one-page loss and all raw outliers remain in the evidence. The investigation stops at this bounded repeat because it exposes no further attributable product cost that would justify another speculative fix or repeated timing runs. Broader performance qualification remains in the existing roadmap.

Conversion quality is qualified independently of timing: nine fresh product/framework lanes passed 44,801 tests with no failures, 18 expected COM opt-in skips and 978 critical passes. Five layout artifacts retain equal geometry and byte-identical PNGs. All 50 Word-native controls passed, and all 55 preview PNGs were personally inspected and match the prior qualified build. Native reference documents were retained rather than regenerated in Word. The existing watermark proportions/image opacity, Cambria Math and general font fidelity, and physical 100-point overflow limits remain.

Both analyses bind their actual reports, producer snapshots, input fingerprints and independent output validation. The original 924 observations and diagnostic exports/profiles resolve through an entry-verified archive after expanded copies were removed; original paths in retained analysis objects are historical references. Archive membership records preserve those paths and hashes. The companion JSON keeps the earlier `90e0` → `43b3` checkpoint and isolated renderer evidence unchanged, with this follow-up in a separate object.

The final follow-up archive is `final-followup-runtime-and-repeat-e01.zip`, 53,684,048 bytes, SHA `142D7957DD258B82754D2549E0BDAB9FF3BFBF8EC2D897CB02D1A04F0D7295AC`. Independent verification checks all 311 decompressed entries and 1,599 original paths, including both exact 37-file runtime graphs and every repeat artifact. `archive-control/membership.json` maps each original path to its content-addressed payload. The original 924-observation cohort and diagnostics remain in `completed-runtime-observations-e01.zip`; quality/native/layout producers remain in `qualified-producers-e01.zip`. The final expanded copies were removed after their consumers completed and archive identities were verified. Visual deliverables and compact evidence remain available separately.

## One-page saving and ARM host follow-up — 2026-10-12

The bounded ARM comparison preserves PDF output but does not qualify stable save latency. It retains 2,080 successful measured operations from 13 of 16 planned launches. A background-load guard stopped the remaining launches; no observations were removed and no rerun replaced the failed qualification. The previous x64 one-page loss remains in the preceding cohort.

This compares the same baseline `43b3c8b4fe7b4c8918155b78741ab7ab73614c9f` and current `e01dcd8f984a54749b72e3e67b36640853722bee`, with all 17 canonical harness files unchanged at `2ed26d029163294b5b2113c979ea59530bd895ff`. No PDF product code changed in this follow-up. A focused PowerForge correction encodes literal tildes in benchmark output paths and replaces the long-path hash separator. This avoids Windows short-name expansion in measured file creation. Both products use the same corrected private net8 runner from owner source `25ee3314c819f2f86d803c17549c51a95b276ff1`; 526 benchmark-service tests pass. Separate sampled profiles found the short-name expansion before the correction and none afterwards. Those profiles include preparation and validation and do not establish how much of the earlier PDF timing difference came from path handling.

The Surface executes native ARM64 PowerShell 7.6.6 on .NET 10.0.12, on AC power. Its two actual shared L2 groups are measured separately: mask `0x3F` has six cores and 12 MiB, and mask `0xFC0` has six cores and 16 MiB. It has no L3 domain. Eleven owned DLL copies were DigiCert-signed for this isolated validation; original runtime graphs remain unchanged, and PE checks preserve the executable payload. Smart App Control remains enabled. The existing Poppler 25.07 x64 validator runs under emulation outside measurement; PDF execution is native ARM64.

The declared comparison uses four balanced rounds, 30 warmups and 40 measurements per case, Normal priority, rotated operations, no removed outliers, and one- and twenty-page direct/copied destinations. The first three rounds provide 1,920 matched measurements; the final 160 baseline-only measurements are retained separately. The table describes only the matched rounds, with 120 measurements per product in each cell. Negative change means lower observed cost. These are unqualified timing observations, not a stable ARM or x64 speed claim.

| L2 group | Operation | Pages | Median ms: baseline → current | Median change | Observed p95 change | Mean allocation change |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| A | `DestinationDirect` | 1 | 10.2848 → 9.9992 | -2.78% | -17.74% | -0.11% |
| A | `DestinationDirect` | 20 | 35.2655 → 33.2812 | -5.63% | -7.41% | -4.93% |
| A | `DestinationCopied` | 1 | 10.3098 → 10.4171 | +1.04% | +9.24% | -0.11% |
| A | `DestinationCopied` | 20 | 34.8417 → 33.4368 | -4.03% | -10.83% | -4.82% |
| B | `DestinationDirect` | 1 | 9.0140 → 8.6980 | -3.51% | +1.59% | -0.13% |
| B | `DestinationDirect` | 20 | 30.2514 → 28.9665 | -4.25% | +16.33% | -4.87% |
| B | `DestinationCopied` | 1 | 8.9596 → 8.8286 | -1.46% | -2.69% | -0.13% |
| B | `DestinationCopied` | 20 | 29.4727 → 28.0768 | -4.74% | +2.85% | -4.77% |

The thirty-second initial load check passed. During the last completed launch, background CPU reached 15.1397% of total capacity, exceeding the declared 15% maximum; a Windows service used about one CPU core at that interval, with Defender also active. The maximum telemetry interval was 1.0564 seconds, so a telemetry gap did not trigger this stop. Monitoring covers whole launches, including untimed preparation and validation; it does not identify an interruption of a particular measured save. All L2-B launches also exceed the declared background p95 limit. The original four-launch repeatability criterion is unmet. Within the three matched rounds, one-page IQR/median exceeds the declared 10% limit, and L2-B distributions are substantially wider. The mixed p95 changes remain visible.

One-page direct-save medians are lower in both groups, while one-page copied saving is 1.04% higher in group A. This does not reproduce a consistent one-page product slowdown or establish a stable improvement. The collected evidence does not justify a speculative PDF change. Stable elapsed-time qualification remains open in the existing roadmap; allocation, elapsed time, CPU, throughput and capacity retain separate claims.

All 2,080 actual measured outputs pass independent provenance, input, exact persisted-PDF byte and full ordered text/glyph-geometry checks. The 112 signed runtime records and 17 harness records also match a fresh post-run check on the Surface. The existing nine-lane, five-layout-artifact and fifty-native-control quality evidence belongs to the unchanged product source; it was not rerun or relabelled as a new-source result. General font fidelity, watermark proportions/image opacity and the physical-margin overflow limits remain unchanged.

The companion JSON stores this separate cohort in `onePageSavingArmFollowup`, including every actual raw sample, matched summaries, unmatched observations, load failures and proof references. The earlier rejected preflight and baseline-only attempt remain separately retained. All 8,917 mapped original files resolve through the entry-verified archive `incomplete-native-arm-producers-and-observations.zip`, 88,392,773 bytes, SHA `81065005182452C73C0FBCE87E45A3D14CC0934B8B6BC5ED0FF2771C9F74E3BA`. Its content-addressed membership preserves 267 decompressed entries, the original unsigned and signed producers, both incomplete attempts and every output. Absolute paths in analysis objects are historical references after expansion cleanup. Earlier benchmark cohorts and their identities remain unchanged.

## Repeated embedded-font saving — 2026-10-12

Reusing an unchanged caller-owned font byte array now avoids repeated private copies and font parsing. In this separate comparison, one-page saving uses 44.83–46.77% less mean managed allocation, and twenty-page saving uses 27.02–28.04% less, across both save paths and CPU cache domains. Every output retains identical PDF bytes and ordered text/glyph geometry. Stable elapsed-time qualification remains open.

The canonical corpus holds one live Carlito byte array but constructs a font family for every save. Previously, each construction copied that array, defeating the existing reference-keyed parsed-font cache. Sampled profiles exposed this repeated parsing cost; source comparison shows it pre-existed both `43b3` and `e01d`, so it does not establish the cause of their earlier timing differences. The new weak cache reuses an owned face snapshot only when bytes and the latest requested family name match exactly. Mutation or a name change replaces the snapshot; earlier families remain immutable, and each document keeps independent glyph usage and subset mappings. Fresh arrays, repeated file reads and alternating family names do not gain this reuse. No public API or production dependency changes.

This cohort compares `e01dcd8f984a54749b72e3e67b36640853722bee` with `4b02eda3a26acaa3b8b35132ad5f2ac81a632fae`. All 17 canonical harness files remain at `2ed26d029163294b5b2113c979ea59530bd895ff`; both products use the same corrected private PowerForge runner. The ordinary AMD Ryzen 9 9950X3D2 run uses native x64 PowerShell 7.6.6/.NET 10.0.12 and freshly queried 96 MiB L3 domains, masks `0xFFFF` and `0xFFFF0000`. Baseline then candidate runs on A; B reverses that order. Four calls retain 640 measurements: 30 warmups, 40 measurements per case, one/twenty pages, direct/copied destinations, Normal priority, rotated operations and no removed outliers. The previously admitted signed baseline copies and unsigned candidate retain their respective producer identities; executable-payload equivalence is recorded separately from signing state.

The table retains every paired cell. Allocation is process-wide managed allocation, including runner overhead; it is not retained heap, peak memory or capacity. Timing columns are descriptive observations from a busy workstation, with one launch per product/domain. Four pre-run activity snapshots observed 38%, 72%, 96% and 81% total CPU use. Lower medians and observed tails do not qualify quiet-host latency or cross-launch repeatability.

| L3 domain | Operation | Pages | Mean allocation MiB: before → after | Allocation change | Observed median ms: before → after | Median change | Observed p95 change |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| A | `DestinationDirect` | 1 | 4.629 → 2.483 | -46.37% | 3.1777 → 2.0114 | -36.70% | -40.88% |
| A | `DestinationDirect` | 20 | 7.769 → 5.613 | -27.75% | 5.3597 → 4.1911 | -21.80% | -16.01% |
| A | `DestinationCopied` | 1 | 4.793 → 2.644 | -44.83% | 3.0790 → 1.9140 | -37.84% | -30.35% |
| A | `DestinationCopied` | 20 | 7.969 → 5.815 | -27.02% | 5.1329 → 4.0118 | -21.84% | -22.50% |
| B | `DestinationDirect` | 1 | 4.630 → 2.465 | -46.77% | 3.2274 → 1.9559 | -39.40% | -14.57% |
| B | `DestinationDirect` | 20 | 7.775 → 5.595 | -28.04% | 5.2592 → 4.0221 | -23.52% | -29.80% |
| B | `DestinationCopied` | 1 | 4.795 → 2.626 | -45.23% | 3.1124 → 1.8504 | -40.55% | -36.53% |
| B | `DestinationCopied` | 20 | 7.973 → 5.797 | -27.29% | 4.9796 → 3.7695 | -24.30% | -33.50% |

The separate Surface attempt produced zero timed operations. Its fresh thirty-second preflight observed total CPU p95 8.2062% and L2-A p95 16.4124%, exceeding the unchanged 8% and 12% limits. The raw native telemetry records background System-process CPU; the underlying cause is unestablished. An earlier controller path error also stopped before timing and remains preserved. Neither attempt was rerun until favorable, and neither contributes samples to the AMD cohort.

Nine fresh source-bound PDF/Word/Excel lanes across net10.0, net8.0 and net472 pass 44,825 tests with zero failures and 18 expected COM opt-in skips. All 1,011 selected critical executions pass; nine TRXs and 39 physical DLL records bind the result to this source. Five final layout artifacts pass geometry and pixel comparison. All 50 Word-native controls pass over 114 complete pages per side; complete PDFs and comparison PNGs equal the prior qualified `e01d` output. Parent inspection covers all 50 composites, including 108 selected visual pages, and all five layout PNGs. Native reference documents were retained, not regenerated in Word. Explicit netstandard2.0 and benchmark net10.0 builds pass. A native ARM smoke comparison also passes both products/save paths at one and twenty pages, eight outputs/84 pages, before the rejected load check. General font fidelity including Cambria Math, watermark proportions/image opacity and the separate 100 pt physical overflow remain open limits.

The companion JSON adds `fontSnapshotSavingFollowup` without changing any prior cohort. It retains all 640 raw samples, eight paired summaries, complete output/provenance proof and both zero-run Surface attempts. Expanded observation paths resolve through `font-snapshot-observations-4b.zip`, 298,023 bytes, SHA `A08BE26239978056A8D3055F336C23C339871C370161ABEE26F70F2B06B4CCF1`: all 57 decompressed entries and 2,499 mapped original files were independently checked. Actual source-bound quality/build producers and original unsigned/signed runtimes resolve through the separately admitted producer archives. No old observations are pooled, overwritten or relabelled. The supported outcome is lower allocation for this finite repeated-font workload; stable one-page elapsed-time evidence remains in the existing roadmap.

## Integrated native ARM saving observations — 2026-10-13

The integrated candidate preserves the selected conversion contract and lower allocation on the finite repeated-font saving workload. All 640 successful observations pass source, producer, PDF-byte, ordered-text and full-glyph-geometry checks. One-page mean managed allocation is 45.51–47.14% lower; twenty-page allocation is 27.36–28.11% lower. Observed median elapsed time is lower in all eight comparisons, but stable latency remains unqualified: background load stopped the planned repeated-round comparison.

This separate cohort compares `e01dcd8f984a54749b72e3e67b36640853722bee` with integrated source `6d7dddc4b5c6008b7748bba613679c17b26e51cd`. The canonical 17 harness files remain at `2ed26d029163294b5b2113c979ea59530bd895ff`; both products use the same corrected private PowerForge runner. Native ARM64 PowerShell 7.6.6/.NET 10.0.12 runs on the Snapdragon X2 Elite Surface. The fresh topology has two six-core L2 domains: A, 12 MiB/mask `0x3F`; B, 16 MiB/mask `0xFC0`. The candidate executes from its exact unsigned graph; the existing signed baseline and runner retain their admitted payload and file identities. Independent x64 Poppler 25.07 runs under emulation outside the operation timer.

The prospective method specifies four balanced rounds, 16 calls and 2,560 samples, with 30 warmups, 40 measurements per case, one/twenty pages, direct/copied destinations, Normal priority, rotated operations and no removed outliers. Four calls completed: baseline then candidate in A, candidate then baseline in B, yielding one launch per product/domain and 640 successful observations. There is no cross-launch repeatability proof. Allocation includes process-wide managed runner overhead and does not measure retained heap, peak memory or capacity. The timing columns below are descriptive under the observed interference; ARM and x64 results remain separate.

| L2 domain | Operation | Pages | Mean allocation MiB: before → after | Allocation change | Observed median ms: before → after | Median change | Observed p95 change |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| A | `DestinationDirect` | 1 | 4.592 → 2.427 | -47.14% | 9.8605 → 6.5778 | -33.29% | -26.21% |
| A | `DestinationDirect` | 20 | 7.732 → 5.559 | -28.11% | 32.8123 → 30.2721 | -7.74% | -10.74% |
| A | `DestinationCopied` | 1 | 4.756 → 2.589 | -45.57% | 10.1055 → 6.3482 | -37.18% | -34.89% |
| A | `DestinationCopied` | 20 | 7.932 → 5.761 | -27.36% | 33.6120 → 29.9346 | -10.94% | -20.69% |
| B | `DestinationDirect` | 1 | 4.603 → 2.436 | -47.08% | 8.6286 → 5.4458 | -36.89% | -18.27% |
| B | `DestinationDirect` | 20 | 7.743 → 5.567 | -28.10% | 30.2373 → 26.2645 | -13.14% | -1.63% |
| B | `DestinationCopied` | 1 | 4.767 → 2.598 | -45.51% | 9.9182 → 6.7163 | -32.28% | -35.85% |
| B | `DestinationCopied` | 20 | 7.944 → 5.770 | -27.37% | 28.5524 → 27.0358 | -5.31% | -9.68% |

The fresh thirty-second preflight passed. During the B baseline launch, a full 1.026-second interval reached 16.3328% conservative native background CPU, above the unchanged 15% maximum. B candidate/baseline background p95 also reached 10.36%/10.44%, above the 8% limit. Defender was the largest readable process contributor at the peak; System and other services were also present. This identifies observed contributors without establishing one underlying cause for every timing difference. The controller stopped further launches and retains the failed load admission.

Two earlier preparation failures remain separate. The first attempt retains 320 successful observations and a failed load guard from a 10.4482-millisecond terminal interval. An observer-only control found zero per-core counter advance in two of eight short intervals, while all eight intervals at the intended one-second cadence were valid. The collector now waits only the remainder of that second before its final snapshot, retaining the child process CPU handle. No operation sample was filtered and no load or stability threshold was loosened. An analysis-only guard correction restores the original positive-to-2.5-second bound for ordinary loop observations: one nominal one-second wait was 0.999625 seconds. The deliberate final observations remain strictly at least one second. The actual full-interval load failure above remains a failure.

The next attempt retains all 160 baseline observations as `Failed`: its longer output directory produced 261-character input paths that Poppler could not open. An untimed control used identical PDF bytes and the same executable: the long path failed with exit 1, while a 68-character path succeeded with exit 0. Only the output root changed to a compact path for the 640-observation cohort. Failed statuses, raw elapsed/allocation observations and artifacts were preserved; none are relabelled as successful or pooled into its results.

The integrated source already passes nine fresh correctness lanes: 45,212 passed, zero failed, 18 expected COM opt-in skips, all 1,017 selected critical executions and 39 physical DLL records. Five layout fixtures pass geometry and pixel comparison. Fifty native-reference controls pass over 114 complete pages; their complete PDFs and PNGs equal the prior qualified candidate. Native Word reference documents were retained, not regenerated. The actual unsigned native ARM smoke adds eight outputs/84 pages with exact prior PDF bytes and ordered glyph geometry. This continuation changes no product or harness code. General font fidelity including Cambria Math, watermark proportions/image opacity and the separate 100 pt physical overflow remain open limits.

The companion JSON adds `integratedUnsignedArmSavingFollowup`; all 20 earlier top-level values remain deeply equal. It retains the 320 successful but unqualified observations, 160 actual failed-validation observations and 640 successful observations as distinct attempts, including all timing/allocation losses and failed admissions. The compact cohort's 2,477 original files are preserved in the independently verified transport archive, SHA `5164C00574F0644E27F158443905530ABD1D73071EB73C5CBE63BF60A9F1C2E2`. Its exact ZIP entry paths resolve historical physical references after cleanup. The earlier AMD, ARM, renderer and full-matrix cohorts remain unchanged. Stable repeated latency across hosts and processor domains remains in the existing roadmap.

## Repeated saving under normal Windows load — 2026-10-13

The complete comparison shows lower saving latency and managed allocation on the measured repeated-font workload. All 32 paired round medians are lower by more than the prospective 3% practical band. One-page pooled medians are 30.64–37.92% lower, p95 is 32.98–42.92% lower, and mean managed allocation is 45.08–46.77% lower. Twenty-page medians are 13.03–23.68% lower and allocation is 27.15–27.97% lower. Twenty-page tail latency remains mixed: domain A copied-save p95 is 9.58% higher, and domain B direct-save p95 is 0.20% higher. Those losses and every raw observation are retained.

This cohort compares product `e01dcd8f984a54749b72e3e67b36640853722bee` with integrated product `6d7dddc4b5c6008b7748bba613679c17b26e51cd`. The source checkout is documentation-only successor `c7bfc9ac4106407e4962a26e730de294f0f93192`; its product, tests and all 17 canonical harness blobs equal the qualified product. Both versions use the same corrected private PowerForge runner and unchanged canonical harness `2ed26d029163294b5b2113c979ea59530bd895ff`. The candidate is its exact unsigned 37-file graph; the previously admitted baseline and runner retain their existing signed bytes. No rebuild or signing took place for this comparison.

Native x64 PowerShell 7.6.6/.NET 10.0.12 runs on the AMD Ryzen 9 9950X3D2. A fresh native cache query confirms two 96 MiB L3 domains with sixteen logical processors each: A/mask `0xFFFF` and B/mask `0xFFFF0000`. The fixed prospective method completes four balanced rounds, 16 launches and 2,560 samples, with 30 warmups and 40 measurements per case, one/twenty pages, direct/copied destinations, Normal priority, rotated operation order and no removed outliers. Product order and cache-domain order alternate between rounds. Independent Poppler 25.07 validation runs outside the measured operation.

| L3 domain | Save destination | Pages | Median ms: before → after | Median change | p95 change | Mean allocation change | Four-launch median range: before / after |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| A | Direct | 1 | 3.6846 → 2.5558 | -30.64% | -33.45% | -46.64% | 11.12% / 4.71% |
| A | Direct | 20 | 6.8172 → 5.9291 | -13.03% | -0.29% | -27.87% | 15.32% / 12.26% |
| A | Copied | 1 | 3.6052 → 2.4699 | -31.49% | -32.98% | -45.08% | 7.79% / 2.11% |
| A | Copied | 20 | 6.5900 → 5.6648 | -14.04% | +9.58% | -27.15% | 8.63% / 5.19% |
| B | Direct | 1 | 3.7016 → 2.3243 | -37.21% | -42.92% | -46.77% | 12.66% / 5.76% |
| B | Direct | 20 | 6.6684 → 5.3813 | -19.30% | +0.20% | -27.97% | 18.77% / 6.44% |
| B | Copied | 1 | 3.6010 → 2.2355 | -37.92% | -37.73% | -45.22% | 10.47% / 6.11% |
| B | Copied | 20 | 6.6320 → 5.0613 | -23.68% | -11.10% | -27.23% | 11.99% / 3.56% |

The comparison deliberately completes under ordinary Windows background activity. All 866 host observations are retained. Total CPU load has a median of 81.08% and p95 of 99.92%; readable external-process CPU has a median of 33.87%. The largest telemetry interval is 1.734 seconds. Native total/domain load includes all activity; readable process attribution can omit exited or inaccessible processes, so these observations do not establish a specific cause for each latency difference. No security settings, CPU affinity policy or power scheme was changed to obtain the result.

The four current-product one-page launch medians span 2.11–6.11% of their median, while the baseline spans 7.79–12.66%. Every paired one-page round is faster: direct A by 26.38–36.32%, copied A by 30.43–34.62%, direct B by 33.87–43.76%, and copied B by 34.68–44.12%. This supports a repeatable direction of improvement on this workload and host. The earlier absolute-stability checks remain visible: all eight cells exceed the prospective 10% per-launch IQR criterion in at least one launch, and some exceed 10% range across launch medians. The cohort therefore makes no quiet-host or tightly bounded absolute-latency claim. Background activity is a reported condition, rather than an execution blocker.

Every one of the 2,560 measured outputs passes independent source/input/producer/settings admission, PDF-byte equality, and full ordered glyph-geometry preservation. The product's existing nine correctness lanes, five layout fixtures and fifty native-reference controls remain bound to the same product source; native Word references were retained, not regenerated. Allocation is process-wide managed allocation including runner overhead, not retained heap or peak memory. These finite warmed-workload results do not establish CPU cost, throughput, capacity, cold-start behavior or a universal speedup. General font fidelity including Cambria Math, watermark proportions/image opacity and the separate physical-overflow case remain open.

The companion JSON adds `integratedNormalLoadAmdSavingFollowup`; all 21 earlier top-level values, including the interrupted ARM comparison, remain deeply equal. The archive retains every original observation, validator artifact, load record and log, plus the zero-timing plan-only serialization failure. That setup failure occurred before measurement; the corrected task script discards the rich plan return object and consumes the shared runner's normalized exports. Earlier slower and failed results remain separate. No sample was filtered or rerun to obtain a favorable result.
