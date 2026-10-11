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

Use a short output root for native Windows Poppler readback; the failed 261-character XHTML path is retained as excluded qualification evidence. Historical binary/TRX copies are preserved in entry-verified archives through the companion's retention mapping. Current profile evidence retains the raw trace and a lossless complete-frame archive.
