# JavaScript analytical-report performance evidence

## Workload and measurement boundary

The measurements cover tagged analytical PDFs and typed XLSX exports through
awaited file sinks. PDF validation preserves the complete logical table, header
associations, identifiers, page content and navigation. Smaller files alone do
not satisfy memory or latency criteria.

The [23 October comparison](#bounded-pdf-and-xlsx-comparison) meets its selected
M8 wide-PDF time and narrow/wide peak-RSS targets for the runtime measured
there. It retains material M9 latency losses, including a higher XLSX median.
Eight M10 comparisons are recorded: [two on 24 October](#m10-comparisons-on-24-october-2026),
the [25 October page-command comparison](#page-command-and-complete-fragment-lifetime),
the [page-local string comparison](#page-local-string-accumulation), the
[physical-row command-parts comparison](#physical-row-command-parts), the
[CID-text/CSV-dispatch comparison](#mapping-local-cid-text-and-primitive-csv-dispatch),
the [numeric-CSV/font-selection comparison](#numeric-csv-and-font-selection-seventh-m10-cohort)
and the [reused-RGB command-text comparison](#reused-rgb-command-text-eighth-m10-cohort).
All except the interrupted second cohort complete 120 independent validations.
None meets all five unchanged M10 criteria. The seventh passes output and both
memory criteria; only its narrow-Helvetica time mean fails the selected time
criterion. The eighth also passes output and both memory criteria, but narrow
Helvetica's time mean and median both rise. Minimal-CSV and XLSX guard losses
also remain visible. Every earlier failed or mixed result stays bound to its
own source; later passes and partial
results do not replace it.

Run the maintained PowerForge suite from the repository root:

```powershell
./Build/BrowserExports/Run-ReportWorkloads.ps1 `
    -BaselineBundle <frozen-baseline.mjs> -OutputRoot <fresh-evidence-root> `
    -Cases tagged-pdf-10000-sink,tagged-wide-inventory-pdf-10000-sink `
    -EquivalentBaselineCases tagged-wide-inventory-pdf-10000-sink `
    -Engine baseline,current -Compare $true -WarmupCount 1 -IterationCount 3
```

The wrapper owns case selection; PowerForge owns rotated ordering, warmups,
measurements and reports. Independent validation reads every exported value and
the tagged structure graph. Validate warmup artifacts as well as measured ones.
Use fresh output roots for changed candidates and retain failed validation,
rejected runs and every measurement. Do not change source, Git provenance or
instrumentation during an accepted cohort.

The frozen baseline is the tagged writer from `d72b7acaeae6c5d68c6a1b33fbd363c355692071`,
with bundle SHA-256
`8de50e9fca18c9efc14f4f266513f7dc90ca0855420273364ad00fc3392ad407`.
This is the actual archived runtime, not a reimplementation of the baseline.

Measurements below use Node 24.18 on an Apple M4 development host with ten
logical processors, normal background activity and OS scheduling. There is no
quiet-host, affinity or forced-GC setting. Peak RSS covers the fresh Node
process, including imported and native memory. Sampled heap and external memory
can miss peaks, overlap other measures, and must not be added to RSS. Sink cases
include awaited regular-file writes; they are not equivalent to Blob retention.

## Historical 23 October 2026 checkpoint

This checkpoint preceded the [integrated source comparison](#integrated-source-comparison).
Its file sizes and memory conclusion describe that earlier candidate.

Source: `07a9c74d23a25e24e825ef5f0706aa0d64e8edd9`.
Runtime SHA-256:
`525db509333e66b57a420fc4da8623d9c8c85c6198100725cf0bbc6fb2f2211f`.
The fixed cohort independently validated twelve measured exports and four
warmups. These are workload observations, not a portable performance ranking.

| Tagged 10,000-row sink workload | Baseline bytes | Current bytes | Mean peak RSS, baseline/current | Median peak RSS change |
| --- | ---: | ---: | ---: | ---: |
| Six columns | 14,098,158 | 3,639,863 | 141.14 / 150.32 MB | +6.85% |
| Twenty columns, horizontal bands | 44,199,021 | 10,748,625 | 279.65 / 272.35 MB | -4.02% |

File sizes fall by 74.18% and 75.68%. Narrow mean peak RSS remains 6.50% higher;
wide mean peak RSS is 2.61% lower. The narrow memory target remains unmet.
Timing varies substantially between launches: narrow medians are 445.01 and
454.41 ms, while means are 441.25 and 1,161.94 ms. Wide medians are 2,205.73
and 1,840.59 ms, while means are 4,624.01 and 1,829.81 ms. Every sample is
retained; these values do not support a universal speedup or regression-free
claim.

## Investigations and disposition

| Change or diagnostic | Evidence and disposition |
| --- | --- |
| Compact tagged structure into bounded object streams | Kept: independent logical graph/content proof passes and file sizes fall. The writer retains a classic fallback when native deflate is initially unavailable. |
| Increase member ceiling from 100 to 200 | Kept within the same 32 KiB decoded-byte ceiling. Prior paired evidence improved the narrow memory result but did not meet its target. Do not repeat this ceiling experiment without a new cause. |
| Reusable decoded arena and numeric batch metadata | Kept: avoids retaining per-member payload arrays. Existing independent byte, xref, UTF-8 and slow-codec tests cover the ownership boundary. Earlier allocation diagnostics are not current RSS acceptance. |
| Reuse ordinary column attributes and avoid scalar-width array copies | Kept after contract review. The 32-output cohort includes CSV/XLSX controls; its narrow PDF RSS remained above baseline. |
| Feed object-stream header and arena body separately | Kept in `07a9c74d`: removes a full payload copy, awaits codec consumption before arena reuse, and passes delayed native-codec proof. The accepted checkpoint above still misses narrow RSS. |
| Allocation sampling including collected objects | Diagnostic only. Its runtime predates `07a9c74d`; total sampled allocation does not identify current retained or peak RSS. M8 already builds the same small structure-object strings, so that string construction is not a newly demonstrated cause. |
| Cohort with a Git commit during execution | Rejected by provenance validation, despite identical runtime bytes. Its timings are excluded. One replacement on the immutable commit produced the accepted checkpoint above. |
| Carry a bounded partial structure batch across page-byte delivery | Kept in `3ccb6153` after structural proof: retained narrow output had 386 object streams for 70,008 members. The 192 recurring 164-member containers occupy only 22.7–24.2 KiB, below the byte ceiling. Page completion forces these partial containers to end. The actual new outputs use 351 narrow and 1,051 wide object streams, reaching the member cap in every full batch. |
| Create subsequent native compressors only when emitting a batch | Kept companion lifetime change in `3ccb6153`. The first capability probe remains eager to preserve classic fallback. Later idle codecs need not overlap page-content compression. Constructor failures after selecting compression must still reject the export. No isolated gain is claimed. |

## Codec-lifetime comparison and remaining target

The combined page-flush/native-constructor candidate at
`3ccb61535ece67d497108fa2891059c738a18951` independently validated sixteen
outputs in one fixed baseline/current cohort. Narrow mean/median peak RSS was
still 3.87%/4.48% above M8; wide mean/median was 1.88%/3.34% lower. Output sizes
were 3,637,087 and 10,728,693 bytes. Narrow sampled heap was lower, while sampled
external memory remained higher. The reduced stream count is proven; it does
not isolate a separate RSS gain for either lifetime change. The narrow memory target remains unmet in that cohort.

## Member-capacity comparison

Candidate `a3c6160c84eca1f96fd76daf6514cffbd19ea75a` raises the member
ceiling from 200 to 256 while retaining the 32 KiB decoded-byte ceiling, arena
size and sink chunks. Actual narrow/wide outputs contain 315/1,025 object
streams, down from 351/1,051; their decoded maxima are 32,766/32,767 bytes.
Both independently retain the same 70,008/210,022 members and 193/386 pages.
The documented member policy is revised explicitly, with index 255 and buffer
reuse covered by the existing compiled-byte fixture.

All sixteen outputs in one fixed M8/current cohort passed independent
validation. Narrow mean/median peak RSS is still 1.62%/1.93% above M8. Wide
mean/median peak RSS is 4.60%/4.00% lower. Outputs are 3,630,112 and 10,723,143
bytes. Narrow engine-time mean/median is 13.71%/14.53% lower; wide mean/median
is 9.85%/40.30% lower. These normal-host results retain all samples and their
launch variation; they do not establish a universal speed or memory claim.
The member-capacity experiment therefore does not meet the narrow-memory
target. Stop increasing the member ceiling; the decoded arena remains 32 KiB.

## Next cause and stop rule

Retained capacity-256 PDFs expose a different source of repeated input. The
six-column report serializes six distinct 46-byte table-attribute dictionaries
10,000 times each; the twenty-column report serializes twenty such dictionaries
10,000 times each. Their 2.76 MB/9.20 MB occupy about 26.84%/27.49% of decoded
object-stream bytes. The existing per-column cache avoids rebuilding the strings,
but each cell still includes and encodes its dictionary.

The next bounded candidate shares immutable ordinary-cell attribute objects
inside their existing table owner. Row-specific headers, spans, IDs, logical
children, structure budgets and physical marked-content ownership must retain
their current meaning. No document-wide interning registry, public option,
provider or dependency is needed. The same semantic representation applies
when compression is disabled or unavailable.

The shared-attribute candidate at
`5c7d39131791269e905ab028714e220b8d5a418c` independently resolves 60,000/200,000
ordinary-cell references to six/twenty dictionaries. The same logical reports
now use 274/821 object streams with decoded maxima of 29,033/31,728 bytes.
Build, 44 focused tests, 502 ordinary tests and one bounded independent
attribute-owner review passed.

All sixteen outputs in its fixed M8/current cohort passed independent
validation. Narrow mean/median peak RSS remains 1.05%/0.93% above M8; wide
mean/median is 5.91%/7.45% lower. Output sizes are 3,576,881/10,514,136 bytes,
74.63%/76.21% below M8. Narrow engine-time mean/median is 177.94%/5.81% higher;
wide is 60.63%/14.86% higher. Every sample remains in the results. Substantial
launch variation under normal background load prevents a general speed claim,
and the narrow-memory target remains unmet.

## Object-location metadata diagnosis

Two actual canonical six-column exports were captured at final awaited byte
delivery with the built-in Node V8 heap snapshot API. These diagnostic children
use synthetic input and an environment whitelist. Snapshots force GC and alter
timing, so their process RSS/timing is excluded from performance acceptance.
Both PDFs passed independent content and structure readback.

On Node 24.18.0 arm64 without pointer compression, M8 and the shared-attribute
candidate each retain a 630,896-byte backing store for ordinary object offsets.
M9 additionally retains an 806,760-byte backing store for compact member
indices. These are measured live-array backing sizes, not a complete
explanation of the peak-RSS difference.

The next bounded candidate stores only compact member indices in lazy typed
chunks. Ordinary offsets and their unwritten sentinel remain unchanged. A
zero slot means an ordinary object; storing index plus one requires Uint16 so
valid index 255 remains distinct. Xref emission must read the chunks directly
without materializing a second whole vector. Preserve large byte offsets,
classic fallback, duplicate/unwritten-object errors and output ownership.

First independently check index 255, chunk crossings, ordinary gaps and full
PDF equivalence, then observe the actual metadata allocation. Only then run
one fixed baseline/current cohort. Do not repeat unchanged candidates for a
favorable result, enlarge the decoded arena or substitute file-size savings
for the narrow-memory target.

## Typed compact-index comparison

Candidate `6cf61fa6b54fb87fad97d75b912aee2ae972304f` stores compact membership
in lazy 4,096-entry Uint16 chunks. Its actual EOF backing storage is 147,816
bytes, down from 806,760 bytes; ordinary offset backing remains 630,896 bytes.
The narrow and wide diagnostic PDFs are byte-identical to the shared-attribute
candidate. Canonical build, 45 focused tests, 503 ordinary tests, independent
artifact readback and one bounded metadata-owner review passed.

All sixteen outputs in one fixed M8/current cohort passed independent
validation. Narrow mean/median peak RSS is still 2.38%/2.31% above M8; wide
mean/median is 10.01%/10.15% lower. Output sizes remain 3,576,881/10,514,136
bytes, 74.63%/76.21% below M8. Narrow engine-time mean is 3.24% higher and its
median is 16.31% lower; wide mean/median is 13.20%/13.27% lower. Sampled
external memory is 68.29%/68.77% higher for narrow and 70.02%/66.43% higher for
wide. Every measured sample and warmup remains in the evidence.

The storage reduction is proven, but it does not establish the narrow-memory
target. This fixed normal-host cohort still misses that target. Stop this
candidate's comparison here: no unchanged rerun for a more favorable outcome.
Further work needs a different measured allocation or lifetime cause. Keep the
32 KiB arena and 256-member ceilings unchanged, and do not infer peak RSS from
the forced-GC EOF snapshots or add sampled heap/external memory to RSS.

## Awaited output-buffer allocation

A separate task-only observer ran the unchanged canonical narrow workload in
both engines, without forced GC or heap sampling. Weak buffer identities at
actual file writes expose 221 distinct 32 KiB backings in the current export:
7,241,728 bytes of cumulative output-buffer allocation for 3,576,881 accepted
PDF bytes. M8 exposes 596 backings, or 19,529,728 bytes. Both PDFs remain
byte-identical to their prior outputs and pass independent readback. The
observer adds overhead; these counts describe allocation traffic, not live
retention or accepted timing/RSS evidence. This cost also exists in M8 and is
not attributed as the cause of the remaining RSS difference.

The next candidate reuses the PDF writer's one 32 KiB output buffer only after
the existing awaited sink acceptance succeeds. The public byte-sink contract
already permits that reuse; Blob and worker destinations copy accepted bytes.
Failed or cancelled writes must leave their pending view stable. No compression
policy, arena/member ceiling, format, API or dependency changes are needed.
Check asynchronous byte integrity and actual backing reuse before one fixed
M8/current cohort. An unchanged retry or diagnostic RSS result cannot satisfy
the narrow-memory target.

The changed writer exposes one 32,768-byte backing across the same 221
acknowledged writes. Its PDF is byte-identical to the independently validated
typed-index output. Forty-six focused and 504 ordinary tests pass, including
delayed byte acceptance and stable pending views after failure/cancellation.
The broad suite also exposed a collector that kept borrowed bytes after
acknowledgement; that fixture now takes its own copy, matching the existing
sink contract. The buffer-reuse allocation reduction is qualified; its effect
on the narrow-memory target still requires the uninstrumented fixed cohort.

Candidate `c303ead6144feb2087a70a9f2e8fa4aad062541f` independently validated
all sixteen outputs in that fixed cohort. Narrow mean/median peak RSS remains
1.93%/1.59% above M8; wide mean/median is 4.62%/2.11% lower. PDF bytes remain
74.63%/76.21% below M8. Narrow engine-time mean/median is 20.17%/5.18% higher;
wide is 41.04%/1.92% higher. Sampled external memory is 37.11%/37.51% higher
for narrow and 23.93%/9.50% higher for wide. All samples remain in the results.
The cumulative buffer-allocation reduction is proven, but the narrow-memory
target is still unmet. No isolated RSS or timing improvement is attributed to
this change, and no unchanged cohort retry is planned.

## Scalar text capture comparison

A current-runtime allocation profile attributes about 78.05 MB of sampled
allocation traffic to the scalar capture call tree, including 23.06 MB beneath
its repeated presentation validation. All three callers already supply owned,
validated, frozen base presentations. A scalar fast path reuses that snapshot
and retains one frozen run, including for an empty string. Rich-run validation,
text budgets and public APIs remain unchanged. Fifty focused and 504 ordinary
tests pass; one bounded capture-owner review accepted the change.

The refreshed profile attributes about 7.48 MB to capture, with no sampled
presentation/color/key/map/join descendants. Its PDF is byte-identical to the
previous independently read current output. These are statistical allocation
observations including collected objects, not retained bytes, peak RSS or
timing acceptance. This removable cost also existed in M8.

Candidate `06b59025e26c29ea766350e18ff4b95fa4c6d39c` passed all sixteen
independent output validations in one fixed normal-host cohort. Narrow
mean/median peak RSS remains 1.28%/0.90% above M8; wide mean/median is
22.66%/8.57% lower. PDF bytes remain 74.63%/76.21% lower. Narrow engine-time
mean is 40.88% lower but its median is 54.65% higher; wide mean/median is
73.89%/66.09% higher. Sampled external memory is 29.92%/37.89% higher for
narrow and 35.83%/30.48% higher for wide. All warmups and measured samples
remain retained; none is removed as an outlier.

The narrow-memory target remains unmet. Do not repeat this unchanged candidate
for a better result or infer its RSS effect from the allocation profiles. A
separate copied-runtime probe found all 467 compression writables unlocked
after reader release, and all 466 first-next-constructor observations unlocked.
That refutes the writer-lock-overlap hypothesis under this diagnostic. It does
not establish native allocation ownership, resource-close timing or a leak.


## Short fitting text

The six-column fixture's 60,000 data cells each fit on a single line. The
unshaped layout path previously created character, advance and owner arrays,
then sliced and joined them even for those fitting cells. Short single-run text
now measures the same font advances with constant scan state and returns the
same line and segment. Empty text, controls, overflowing text, shaped text,
multiple runs and long text retain their existing paths and cooperative checks.
The `.001` fit tolerance and first-error order remain unchanged.

The current profile attributes 80.03 MB of sampled allocation traffic to line
wrapping, compared with 180.72 MB before this change. Result objects remain;
the sample does not establish retained memory or RSS savings. All 505 ordinary
and 187 focused PDF tests pass. Ten public-API PDF profiles, including embedded
fonts, links, controls and fit boundaries, are byte-identical to the prior
runtime and independently pass qpdf/pypdf. Both full workload PDFs also retain
their exact validated bytes. One bounded line-layout review accepted the change.

Candidate `42939bde2230e808edec0755242b9f629f84cc1d` passed all sixteen
independent output validations in one fixed normal-host cohort. Narrow peak
RSS remains 0.27% higher by mean and 0.50% higher by median than M8. Wide peak
RSS is 3.56% higher by mean and 7.22% lower by median. PDF bytes remain
74.63%/76.21% lower for narrow/wide. Narrow engine time is 7.74% higher by mean
and 26.85% lower by median; wide mean/median is 21.41%/30.42% lower. Sampled
external memory is 39.13%/37.11% higher for narrow and 29.73%/18.52% higher for
wide. Every warmup and measured sample is retained, including the wide sample
with 328,941,568 bytes peak RSS. No outlier was removed.

The narrow-memory target remains unmet in this cohort. This candidate is not
rerun for a more favorable result. Allocation traffic, heap, external memory
and process RSS remain distinct observations; these measurements establish no
universal speedup, no-regression or constant-memory claim.


## Lazy link tracking

The preceding runtime's `element()` frame accounts for 5.61 MB of sampled
self allocation. The independently validated six-column document contains
70,009 logical elements and no links. Each element nevertheless created an
empty link-identity array. The array is now allocated only when a first link
identity is registered for that logical owner. Deduplication, completion and
parent cleanup retain their existing lifecycle. No byte cost or RSS saving is
assigned to this property from the profile alone.

All 505 ordinary and 187 focused PDF tests pass. Ten layout profiles and six
caption/group semantic fixtures retain identical before/after PDF bytes and
pass independent readback. Narrow and wide 10,000-row workload bytes also
remain identical. The mechanical optional-allocation change was checked
proportionately; the prior semantic-owner reviews remain applicable.

Candidate `9905165f0eb061ac329f5fc4e3531e5c5895923a` passed all sixteen
independent output validations in one fixed normal-host cohort. Narrow mean
peak RSS remains 0.11% above M8; its median is 1.54% lower. Wide mean/median
peak RSS is 4.40%/3.69% lower. Narrow engine-time mean/median is 5.00%/33.49%
higher, while wide mean/median is 13.44%/10.93% lower. Sampled external memory
is 33.22%/34.68% higher for narrow and 15.96%/10.93% higher for wide. PDF sizes
remain 74.63%/76.21% lower. All samples and warmups are retained.

The narrow target requires both mean and median, so it remains unmet in this
cohort. The candidate is not repeated to seek a different outcome, and the
measured differences do not establish an isolated effect for the optional list.


## Single-run font selection

The short single-run text path selects its font once and defers the temporary
face array until general layout needs it. Multi-run selection, fallback
selection order and text geometry retain their existing behavior. All 505
ordinary and 187 focused PDF tests pass; twelve layout profiles and six
semantic fixtures retain identical PDF bytes and independent readback.

Candidate `6d58b2763f43b636db327e4aee3e770563bca133` passed all sixteen
independent output validations in one fixed normal-host cohort. Narrow mean
peak RSS is 1.31% above M8, while its median is 0.20% lower. Wide mean/median
peak RSS is 20.15%/11.22% lower. Narrow engine-time mean/median is
22.68%/26.99% lower; wide mean/median is 25.80%/24.42% lower. PDF sizes remain
74.63%/76.21% lower. Sampled external memory is 45.12%/46.05% higher for
narrow and 34.14%/48.94% higher for wide. Every sample is retained.

Both narrow mean and median must meet the target, so this cohort remains
unmet. It is not repeated for a more favorable outcome. These measurements
do not isolate an RSS effect for the face array or establish a universal
speedup, no-regression or constant-memory claim.


## Same-page marked content

PDF structure elements now identify same-page marked content with integer
children and an explicit page on their logical owner. Later-page fragments and
annotation references retain explicit page references. The six-column workload
replaces 60,006 per-mark reference strings while preserving its 70,009 logical
elements; the wide workload preserves 210,023 elements and 200,020 marks.
These counts do not establish retained-memory or RSS savings.

All 510 ordinary and 192 focused PDF tests pass. Nine independent-reader oracle
cases accept both encodings and reject incorrect page, header and owner
bindings. Twenty before/after artifact pairs preserve decoded page operators,
fonts, physical text and link geometry across 633 pages. Existing independent
readers verify every narrow and wide value, association and ParentTree owner,
plus six mixed-page caption/group fixtures. A bounded independent review
accepted the private encoding and its receiving inspectors.

Candidate `ddd63a278856ae161fcd76cc6044ed73e1dba11c` passes all sixteen
output validations in one fixed normal-host cohort. Narrow mean/median peak
RSS is 0.36%/0.50% above M8; wide mean/median is 18.51%/6.42% lower. Narrow
engine-time mean/median is 13.11%/16.30% lower, and wide mean/median is
19.85%/18.69% lower. Output sizes are 74.71%/77.02% lower. Sampled external
memory is 47.19%/47.61% higher for narrow and 20.93%/18.46% higher for wide.
All samples remain included. Both narrow RSS criteria remain unmet; this
candidate is not repeated to seek a different result. No isolated RSS effect
or universal performance guarantee follows from these observations.

## Buffered structure completion

Completed structure leaves can append synchronously to an existing compact
batch when no flush is needed. Fresh batches, UTF-8 overflow and the exact
256-member boundary retain the awaited path. Partial child completion keeps
the unfinished parent open and skips children already stored. Public APIs,
native compression and serialized bytes remain unchanged.

All 512 ordinary and 194 focused PDF tests pass. Twelve geometry profiles,
six semantic fixtures and both full workloads retain their exact independently
validated bytes. A bounded independent ownership review accepted the change.
Sampled allocation traffic beneath completion falls from 155.17 MB to
83.55 MB, including the new buffered helpers; this measures allocation traffic,
not retained memory, native memory or process RSS.

Candidate `12455d1b1a84689a22f967585e1477368bf26638` passes all sixteen
output validations in one fixed normal-host cohort. Narrow mean/median peak
RSS remains 1.75%/3.96% above M8; wide mean/median is 9.04%/10.69% lower.
Narrow engine-time mean/median is 32.61%/9.92% lower, and wide mean/median is
3.25%/29.21% lower. Output sizes remain 74.71%/77.02% lower. Sampled external
memory is 40.72%/48.52% higher for narrow and 86.38%/89.07% higher for wide.
All warmups and measured samples remain included.

Both narrow RSS criteria remain unmet. This candidate is not repeated to seek
a different result. The next bounded investigation concerns the extra input
stream and pipe-through transport around the existing native compressor;
mandatory codec work remains. No isolated RSS effect or universal performance
guarantee follows from the observations above.

## Native compressor transport prototype

A local prototype replaced the compressor's artificial input stream and
pipe-through transport with a concurrent writer and reader. Its first-error,
cancellation and buffer-ownership checks passed, as did 517 ordinary and 199
focused PDF tests and exact artifact comparisons. Independent review found no
actionable correctness issue in that boundary.

The complete sampled owner scope, including detached input continuations,
increased from 9.87 MB to 12.07 MB once its new lifetime and cancellation work
was included. Mandatory native codec construction remained separate. The
current profile also contains 499,648 bytes with unresolved ancestry; they are
not assigned to any owner. This does not prove an RSS regression, but it gives
no basis for retaining the extra transport complexity. The prototype was
removed before an RSS comparison. Its patch, diagnostics and disposition are
retained; the original compressor transport and its tests remain in use.

## Standard-font byte encoding

The standard-font encoder now reuses a private 256-entry two-digit hexadecimal
lookup. It still validates each scalar and uses the same WinAnsi mapping;
embedded-font CID conversion is unchanged. One public PDF fixture checks all
217 supported byte values and the four standard faces. All 513 ordinary and
195 focused PDF tests pass; twelve layout profiles, six semantic fixtures and
both full workloads preserve exact independently validated bytes.

In the prototype capture, the unchanged selected font method's inclusive
sampled traffic fell from 26.86 MB to 16.56 MB. Its scoped per-byte `toString`
traffic fell from 10.24 MB to zero observed. Lookup initialization was not
sampled, which does not mean its cost is zero. The font method remains identical
after removing the transport prototype. These observations justify the bounded
font change, not an RSS claim. The selected candidate still needs its one fixed
M8 comparison; both narrow RSS criteria remain required.

Candidate `768e1483117dcbf31194c67887c9923403f1eed2`, containing the font
lookup with the original compressor transport, passes all sixteen output
validations in one fixed normal-host cohort. Narrow mean/median peak RSS is
2.71%/2.20% above M8; wide mean/median is 10.05%/8.89% lower. Narrow engine-time
mean/median is 37.15%/28.13% lower, and wide mean/median is 40.81%/10.96% lower.
Output sizes remain 74.71%/77.02% lower. Sampled external memory is
49.96%/52.42% higher for narrow and 77.85%/79.74% higher for wide. Every warmup
and measured sample remains included.

Both narrow RSS criteria remain unmet, and the candidate is not repeated.
The next bounded cause is repeated ordinary-cell default presentation capture;
explicit cell styles and special table rows must retain their current path.
No isolated RSS effect or universal performance claim follows from this cohort.

## Table-local default presentations

Ordinary cells without an authored presentation now reuse a table-local frozen
default for their column and plain or alternate row. Defaults are created only
when reached, at most twice per column. Explicit cell styles, headings, spans,
footers and generated summaries retain their full merge path. Format callbacks,
captured getters, lazy validation, links, geometry and cancellation are unchanged.

All 516 ordinary and 198 focused PDF tests pass; the final five input-contract
tests also pass after refining a cancellation assertion to check no later writes.
Twelve geometry profiles, six semantic fixtures and both full workloads preserve
exact independently validated bytes. One fresh independent ownership review
accepted this private sharing boundary without actionable findings.

Comparing the historical buffered-completion profile with the current capture,
sampled cell self traffic falls from 80.65 MB to 66.84 MB. The deduplicated
cell/data/construction scope falls from 206.25 MB to 184.25 MB, including sampled
new initialization costs. Presentation's inclusive subtree instead rises from
8.30 MB to 9.41 MB. These overlapping sampled scopes cannot establish how much
isolated saving came from reuse; the font change also exists outside this owner.
They are allocation-traffic diagnostics, not retained-memory or RSS measurements.
Both narrow RSS statistics still require the one fixed M8 comparison.

Candidate `ff279b2b5e5bfd94d7cf84005a5f3dcd4d8764cb` passes all sixteen
output validations in one fixed normal-host cohort. Narrow mean/median peak
RSS is 2.84%/3.16% above M8; wide mean/median is 4.59%/3.58% lower. Narrow
engine-time mean/median is 22.45%/20.17% lower, and wide mean/median is
28.44%/31.61% lower. Output sizes remain 74.71%/77.02% lower. Sampled heap
is 5.24%/5.95% lower for narrow, while sampled external memory is
49.67%/48.83% higher. These memory layers are separate and not additive.
Every warmup and measured sample remains included. Both narrow RSS criteria
remain unmet, and the candidate is not repeated to seek a different result.
No isolated RSS attribution or universal performance claim follows.

## Raw structure containers after the native probe

The first bounded structure container retains the existing compressed capability
probe. Later containers write their existing index and body directly with a known
length instead of constructing another native compressor. Object/xref streams
permit an absent compression filter. The 256-member and 32 KiB decoded limits,
plain fallback before compaction, page/font/image codecs and awaited sink ownership
are unchanged. Raw bytes remain borrowed until acceptance, before arena reuse.

All 517 ordinary and 199 focused PDF tests pass, including seventeen object-stream
contracts. Independent qpdf/pypdf validation covers both full workloads; six
semantic fixtures and twelve layout profiles pass. The actual built runtime
produces the same narrow/wide hashes as the isolated serialization prototype.
Decoded physical commands and font bindings remain identical to the preceding
compressed-container candidate across all 193 and 386 pages.

The test reader previously mistook two raw container payloads containing
`/Type /StructElem` text for actual structure dictionaries. The shared eligibility
selector now excludes streams and checks the outer structure dictionary. Its
strict child-token parser and all expected structure/ownership graphs remain
unchanged. The original failure and independent valid-PDF proof are retained.

In the two instrumented diagnostic exports, native compressor constructions fell
from the preceding source/artifact-derived 467/1,207 to 194/387. Every diagnostic
codec group and producer closed. This establishes fewer constructions, not native
C++ deallocation or an isolated RSS saving. The resulting files are 9,163,300 and
28,185,805 bytes: larger than compressing every container, but still 35.00% and
36.23% smaller than M8. This explicit file-size/codec-cost tradeoff requires its
one fixed normal-host comparison; both narrow peak-RSS statistics remain required.

Candidate `d67c7ee00ef62ac21f7f2141b1f07cdedce70b05` passes all sixteen
output validations in one fixed normal-host cohort. Narrow mean/median peak RSS
is 1.92%/1.40% below M8; wide mean/median is 10.37%/8.15% below M8. Narrow
engine-time mean/median is 42.43%/46.14% lower, and wide mean/median is
31.25%/25.55% lower. Sampled narrow heap is 8.47%/7.17% lower and sampled
external memory is 11.74%/12.43% lower. Wide sampled heap is 10.49%/9.35%
lower, while its external memory is 7.89%/6.14% higher. Both narrow RSS
criteria are met in this cohort. Every warmup and measured sample is retained;
one normal-host n=3 comparison does not establish a universal performance claim.

The subsequent integration retains the independently merged PNG validation,
Node support and WinAnsi test-reader changes. PNG code changes the bundle but
is not called by either image-free workload. The integrated source receives one
final fixed comparison, with the same M8 baseline and policy, before its package
and browser acceptance. The preceding result remains bound to its exact source.

## Integrated source comparison

The integrated candidate `66533554d1fa0a7b7623c1521b6b3b3f9a626666`
passes all sixteen validations in its one fixed normal-host cohort, including the
same exact narrow/wide PDF hashes. Narrow mean/median peak RSS is 0.88%/0.50%
below M8; wide is 10.61%/11.67% below. Both narrow RSS criteria are met.
Narrow sampled heap is 9.10%/8.74% lower and sampled external memory is
2.05%/2.18% lower. Wide sampled heap is 12.71%/10.34% lower, while its sampled
external memory is 3.20%/7.37% higher. These layers remain separate and nonadditive.

Timing does not show a consistent speedup across launches: this integrated
cohort's narrow engine-time mean/median is 60.69%/26.04% higher than M8; wide
mean is 31.55% lower and median is 0.79% higher. These losses remain included,
as do all four warmups and twelve measured samples. The earlier source's faster
cohort is not substituted for this result. Identical output and an unused PNG
path do not prove the timing difference is noise or establish its cause. No
universal speedup, performance-no-regression or constant-memory claim is made.

## Export-contract repair comparison

Source `a64cf0aa09ec54bb3ddb5998594c3ff47bc22f73`, runtime
`5538e04884097a6a4033aa3337f5e8adf8a0e1deadbd7a9922f5bf3a561a2959`,
passes all sixteen output validations in its one fixed normal-host cohort.
Narrow mean/median peak RSS is 2.93%/5.38% above M8, so both criteria remain
unmet for this source. Wide mean/median peak RSS is 8.30%/11.32% below M8.
Narrow engine-time mean/median is 43.03%/34.98% lower; wide is 34.80%/31.94%
lower. Output hashes and sizes remain unchanged from the integrated comparison.
All four warmups and twelve measured samples are retained.

These cases use an explicit borrowing byte sink; the repaired native
`WritableStream` snapshot branch is not reached. The rendered-text budget helper
returns the original text because these cases do not set `maxTextCharacters`.
The comparison does not attribute its RSS result to either repair. Sampled heap
and external maxima do not decompose the RSS difference, and background load
alone does not establish a cause. This failed cohort is not repeated unchanged.

## Lazy inline ownership comparison

Scalar text lines no longer allocate an empty ownership map. The map is created
only in the existing tagged `actualText` path, which retains logical-order marked
content and links before visual-order painting. This is a removable common cost,
not an M9-introduced regression: the earlier emitted method is byte-identical
across M8, the retained table-presentation profile and the repaired source.

Those actual profiles sampled 11.56 MB and 11.98 MB of direct map allocation
traffic beneath `inline`, including collected objects. A separate diagnostic
using the canonical narrow input counts 61,158 unpopulated maps before the change
and zero afterward. Both diagnostic PDFs remain byte-for-byte identical to the
independently validated 9,163,300-byte workload. Instrumentation exists only in
private copies, and its timings or memory are not acceptance measurements.
All 524 ordinary and 65 focused shaping, structure, table, band, grouped,
streaming and document tests pass.

Source `a547cbb0cbb3e6b85baad21d530b4dae144e8564`, runtime
`5cb3e153d6de0821e6f4a517701c7e4a0107aa05e8ecb20d725daf8a3397fbab`,
passes all sixteen output validations in one new fixed normal-host cohort.

| Workload | Peak RSS mean / median vs M8 | Engine time mean / median vs M8 | File size vs M8 |
| --- | --- | --- | --- |
| Six-column tagged, 10,000 rows | 1.28% / 2.11% lower | 34.02% / 31.47% lower | 35.00% smaller |
| Twenty-column tagged, 10,000 rows | 14.85% / 14.43% lower | 5.10% / 1.98% higher | 36.23% smaller |

Both narrow peak-RSS criteria are met in this cohort. Narrow sampled heap
mean/median is 7.51%/10.44% lower; sampled external memory is 10.14%/0.14%
higher. Wide sampled heap is 13.38%/12.37% lower, while sampled external memory
is 25.20%/24.09% higher. These memory layers remain separate and nonadditive.
The files retain the exact earlier hashes, at 9,163,300 and 28,185,805 bytes.

Every sample remains included, including the wide engine-time loss. The preceding
failed source and earlier faster cohorts are not substituted or pooled. The
allocation-count change does not isolate its contribution to RSS; one normal-host
n=3 comparison does not establish a universal speedup, performance-no-regression
or constant-memory guarantee.

## Bounded PDF and XLSX comparison

Measured 23 October 2026 UTC.

One fixed PowerForge cohort compares the committed M8 and M9 runtimes with the current measured runtime: three workloads × three runtimes × one warmup and three measurements. All 36 actual outputs passed independent readback. This is a normal shared Apple M4 development host, macOS Arm64 and Node 24.18.0, with rotated ordering, no affinity or priority changes, and no outlier removal. Host-load observations and every sample remain in the record.

Each sample imports its module in a fresh process before timing; its JIT state is fresh. Lazy row construction and awaited regular-file sink writes are included. OS page cache applies. Peak RSS covers that Node process through operation completion, including imports and native memory. These results do not describe persistent warm exports, retained Blob/buffer delivery, durable file flushes or constant memory.

Values below are mean / median; timings are milliseconds and RSS is MiB. Rounding is for display.

| Workload | Runtime | Export ms: mean / median | Peak RSS MiB: mean / median |
|---|---|---:|---:|
| Tagged PDF, 10,000 × 6 | M8 | 883.70 / 858.30 | 134.77 / 134.38 |
| Tagged PDF, 10,000 × 6 | M9 | 977.71 / 806.43 | 129.11 / 130.97 |
| Tagged PDF, 10,000 × 6 | Current | 1,824.44 / 860.31 | 128.21 / 129.67 |
| Wide tagged PDF, 10,000 × 20 | M8 | 6,843.80 / 5,503.76 | 253.83 / 251.33 |
| Wide tagged PDF, 10,000 × 20 | M9 | 2,370.94 / 1,131.97 | 236.07 / 238.03 |
| Wide tagged PDF, 10,000 × 20 | Current | 3,924.04 / 4,105.39 | 229.37 / 227.47 |
| Typed XLSX, 100,000 × 6 | M8 | 1,986.29 / 1,888.75 | 81.68 / 82.55 |
| Typed XLSX, 100,000 × 6 | M9 | 1,676.31 / 1,367.55 | 81.04 / 81.16 |
| Typed XLSX, 100,000 × 6 | Current | 1,247.75 / 1,518.41 | 78.74 / 78.89 |

The three frozen M8 checks pass in this run: narrow-PDF RSS mean and median are at or below M8; wide-PDF RSS mean and median are lower; and wide-PDF time mean and median are lower. That does not establish no regression against M9. The following material M9 latency losses remain **unwaived**:

- Narrow PDF time: **+86.60% mean, +6.68% median**. Its current raw times are 378.71, 4,234.30 and 860.31 ms; the long sample is retained. Against M8, narrow time is +106.45% mean and +0.23% median.
- Wide PDF time: **+65.51% mean, +262.68% median**. Its current raw times are 5,957.99, 4,105.39 and 1,708.75 ms.
- XLSX time: mean is 25.57% lower than M9, but **median is 11.03% higher**. Both results are retained rather than selecting the favorable statistic.

Current RSS mean/median is below M9 for all three workloads: narrow PDF by 0.70% / 0.99%, wide PDF by 2.84% / 4.44%, and XLSX by 2.84% / 2.79%. The tagged PDFs retain the same bytes as M9: 9,163,300 bytes for the narrow report and 28,185,805 bytes for the wide report. Smaller files or lower memory do not waive latency losses.

An earlier integrated runtime, `d265fdea78c04366a5088a0425cbb2658905a8596135b87df9336e34bb81b6a7`, remains separate evidence. Its 24 PDF outputs passed, but narrow time was 1,117.39 / 359.89 ms against M9 523.98 / 314.21 ms, including a retained 2,694.18-ms sample; wide mean RSS was 0.35% above M9. A separate 100,000 × 6 XLSX trial measured 3,833.95 / 3,797.73 ms against M9 3,311.41 / 3,645.82 ms. CRC loop unrolling did not earn its predeclared full-workload retention criterion and was reverted. This changed runtime received the new 36-output cohort above; the earlier losses were not erased by an unchanged rerun.

One bounded narrow-PDF trace of that earlier runtime produced byte-identical, independently readable PDFs with the same 397 sink writes and 582 compressed reads. It did not establish the cause of the original 2,694-ms sample or a new production hotspot. Instrumented timing/RSS is diagnostic only and does not replace acceptance samples; CPU samples, phase durations and GC intervals are not additive. This trace does not qualify the later runtime.

One separate wide-PDF diagnostic pair also produced byte-identical, independently readable output with matching 1,088 sink writes and 1,161 compressed reads. That perturbed pair did not reproduce a current-only bottleneck or establish the cause of the fixed-cohort latency difference. It provides no acceptance timing or RSS result, does not replace the 36-output cohort, and does not waive its M9 losses.

For reproduction, use the committed `OfficeIMO.JavaScript/bundles/officeimo.mjs` bytes for M8 and M9 without rebuilding them, and verify the hashes before running the existing canonical wrapper:

| Runtime | Revision | `officeimo.mjs` SHA-256 |
|---|---|---|
| M8 | `d72b7acaeae6c5d68c6a1b33fbd363c355692071` | `8de50e9fca18c9efc14f4f266513f7dc90ca0855420273364ad00fc3392ad407` |
| M9 | `7f27e4dfe28a8e61c5042a730778bb0d9bcd8990` | `5cb3e153d6de0821e6f4a517701c7e4a0107aa05e8ecb20d725daf8a3397fbab` |
| Current measured bytes | Source identified by hash | `efc8e4d16d24d19a4df3b023a665123a198476ab03aefba5b9c7e4915ff8a99b` |

```powershell
$refs = @{ m8 = './baseline/m8.mjs'; m9 = './baseline/m9.mjs' }
./Build/BrowserExports/Run-ReportWorkloads.ps1 `
    -BaselineBundles $refs -OutputRoot ./Ignore/Benchmarks/report-sinks-20261023 `
    -Cases tagged-pdf-10000-sink,tagged-wide-inventory-pdf-10000-sink,inventory-xlsx-100000-sink `
    -EquivalentBaselineCases tagged-wide-inventory-pdf-10000-sink `
    -Engine m8,m9,current -WarmupCount 1 -IterationCount 3
```

Run identity: `20261023-220212-951eba5a`. The shared runner and independent readers own scheduling and validation; keep all raw samples and losses. This finite cohort provides no universal speedup, no universal no-regression result and no quiet-machine or cross-host guarantee.

## M10 comparisons on 24 October 2026

Neither candidate below meets the fixed M10 criteria. The first cohort completes
all 120 independent validations but fails PDF time and peak-RSS criteria. The
second is interrupted by the whole-command deadline; its completed lanes already
fail criteria, so completing the missing embedded-font samples alone cannot
qualify that candidate.

Both use the existing PowerForge runner on the normal shared Apple M4
development host, macOS Arm64 and Node 24.18.0. Each selected case/runtime has
one warmup and five measured exports, rotated ordering, no outlier removal and
no affinity, priority or forced-GC change. The fixed matrix has 120 actual
outputs: 48 Helvetica PDF, 48 CSV/XLSX and 24 embedded-Carlito PDF outputs.
Helvetica and Carlito remain separate profiles.

Each export imports its module before the operation timer in a fresh Node
process; JIT state remains fresh. Lazy input construction and awaited regular-file
sink writes are included. Embedded-font setup precedes the timer. Lifetime peak
RSS includes imports, font setup and native allocations; sampled heap/external
maxima do not decompose it and must not be added to RSS. `EngineMs` is the inner
export metric, not the shared runner's outer `durationMs`. First delivery means
the first accepted sink bytes, not a usable first page or record.

The exact runtime bindings are:

| Role | Source or package | `officeimo.mjs` SHA-256 |
| --- | --- | --- |
| Public baseline | `@evotecit/officeimo` 0.1.2 | `efc8e4d16d24d19a4df3b023a665123a198476ab03aefba5b9c7e4915ff8a99b` |
| First M10 candidate | `3f7553503c7b25f664a512374a89535aac55459c` | `2c71940deb30e58f1964cec9f375534a189533333902999c867d0ed5011b4928` |
| Revised M10 candidate | `213a70fbdc96ea9d698419b822b9b55fe7f9c2b4` | `7007b569d584e8ca39dbb99267b0f290eab629fcc389f9dc2b1d631be1dbab35` |

The M8 and M9 baselines retain the exact revisions and bundle hashes in the
preceding reproduction table. They are reused as frozen compiled bytes without
rebuilding. The revised M10 runtime withdraws a per-call projection map, restores
`for...of` traversal and avoids redundant lookups in existing font caches. That
source change justifies a separate cohort; it does not establish an isolated
speed or memory contribution.

The five criteria remain unchanged:

1. Independently validate all 120 actual outputs, including warmups.
2. Retain exact public-0.1.2 output bytes for the changed PDF/CSV cases. XLSX is
   independently checked for typed values and schema; timestamp-dependent ZIP
   identity is reported separately.
3. Each of the four PDF and two quote-all CSV cases must have lower export-time
   **mean and median** than public 0.1.2.
4. Each PDF case must have peak-RSS **mean and median** no higher than public
   0.1.2.
5. Both original Helvetica PDF cases must preserve peak-RSS **mean and median**
   no higher than M8.

### First M10 cohort: complete, criteria unmet

All 20 warmups and 100 measured outputs pass independent validation. Every
changed PDF/CSV output is byte-identical to its public-0.1.2 counterpart. This
correctness result does not satisfy the time or memory criteria.

Each value below is **mean / median**. The arrows show **public 0.1.2 → first
candidate**. Times are milliseconds; RSS is MiB. Percentage changes use
unrounded measurements; display values are rounded.

| Workload | Export ms, public → candidate | Peak RSS MiB, public → candidate | Time change | RSS change |
| --- | ---: | ---: | ---: | ---: |
| Helvetica tagged PDF, 10,000 × 6 | 842.23 / 711.97 → 1,060.28 / 1,032.25 | 131.85 / 131.91 → 131.64 / 132.70 | +25.89% / +44.99% | -0.16% / +0.60% |
| Helvetica tagged PDF, 10,000 × 20 | 3,455.10 / 3,139.47 → 3,448.32 / 3,082.86 | 234.72 / 241.38 → 245.79 / 234.64 | -0.20% / -1.80% | +4.72% / -2.79% |
| Carlito tagged PDF, 10,000 × 6 | 932.17 / 983.34 → 1,098.52 / 919.75 | 134.65 / 133.95 → 136.29 / 135.86 | +17.84% / -6.47% | +1.22% / +1.42% |
| Carlito tagged PDF, 10,000 × 20 | 2,393.07 / 3,036.86 → 2,140.11 / 1,642.40 | 236.50 / 236.05 → 242.22 / 243.36 | -10.57% / -45.92% | +2.42% / +3.10% |
| Minimal CSV, 100,000 × 6 | 267.71 / 268.14 → 270.00 / 255.19 | 74.08 / 74.02 → 74.36 / 74.31 | +0.86% / -4.83% | +0.39% / +0.40% |
| Quote-all CSV, 100,000 × 6 | 284.84 / 290.55 → 242.74 / 258.97 | 74.79 / 74.81 → 74.42 / 74.52 | -14.78% / -10.87% | -0.50% / -0.40% |
| Quote-all CSV, 100,000 × 20 | 688.75 / 768.49 → 636.19 / 660.19 | 80.70 / 80.11 → 75.42 / 75.31 | -7.63% / -14.09% | -6.54% / -5.99% |
| Typed XLSX, 100,000 × 6 | 1,244.43 / 1,236.19 → 1,147.85 / 1,100.57 | 79.63 / 79.12 → 80.16 / 80.14 | -7.76% / -10.97% | +0.67% / +1.28% |

The PDF time criterion fails for narrow Helvetica and the Carlito narrow mean.
The public-baseline RSS criterion fails for the Helvetica narrow median, the
Helvetica wide mean, and both statistics of each Carlito case. Original M8
Helvetica RSS gains remain in this cohort: narrow mean/median is 1.43% / 1.89%
lower and wide is 5.34% / 6.99% lower. That historical-baseline result does not
waive the newer public-baseline or time criteria. Minimal-CSV mean time and
XLSX memory losses remain visible as guard observations.

Canonical run identities:

| Stage | Run identity | Independently validated warmups / measured |
| --- | --- | ---: |
| Helvetica PDF | `20261024-121852-60e05b25` | 8 / 40 |
| CSV/XLSX | `20261024-124832-60e80484` | 8 / 40 |
| Embedded Carlito PDF | `20261024-125023-246f326f` | 4 / 20 |

Each run retains its canonical `samples.json`, `summary.json`, `comparison.json`
and `comparison.md`, plus the per-iteration producer measurements and independent
validation receipts. The first candidate's runtime and packages remain retained
with their hashes, so its manifests are not reinterpreted through a subsequently
rebuilt worktree path.

### Revised M10 cohort: interrupted, completed criteria still unmet

The whole-command supervisor has a 3,600-second deadline. Its receipt ends at
14:27:53.308791 UTC after 3,619.75 seconds with a timeout and exit code -9.
This establishes an incomplete cohort; it does not establish an export crash or
the cause of any measured latency difference.

The Helvetica and CSV/XLSX stages validate all 96 outputs. Embedded Carlito has
five independently validated measured outputs, four produced but unvalidated
warmups, one produced but unvalidated measured output, and fourteen missing
measured outputs. Totals are **101 validated, 5 produced but unvalidated and 14
missing**. The embedded stage has no terminal canonical summary, comparison,
sample aggregate or final runtime-verification receipt. Its partial samples are
retained but are not used for five-sample statistics.

The completed five-sample cases are shown below. Arrows now mean **public 0.1.2
→ revised candidate**; units and mean/median ordering are unchanged.

| Workload | Export ms, public → candidate | Peak RSS MiB, public → candidate | Time change | RSS change |
| --- | ---: | ---: | ---: | ---: |
| Helvetica tagged PDF, 10,000 × 6 | 1,371.05 / 1,215.09 → 1,114.39 / 1,112.60 | 130.17 / 129.98 → 131.54 / 132.06 | -18.72% / -8.43% | +1.06% / +1.60% |
| Helvetica tagged PDF, 10,000 × 20 | 6,534.79 / 3,091.82 → 7,790.75 / 3,465.32 | 227.97 / 231.52 → 225.43 / 227.84 | +19.22% / +12.08% | -1.11% / -1.59% |
| Minimal CSV, 100,000 × 6 | 1,026.60 / 572.66 → 1,024.03 / 508.64 | 73.87 / 73.81 → 72.64 / 73.73 | -0.25% / -11.18% | -1.66% / -0.11% |
| Quote-all CSV, 100,000 × 6 | 989.06 / 637.20 → 558.02 / 507.99 | 73.16 / 74.42 → 74.43 / 74.61 | -43.58% / -20.28% | +1.74% / +0.25% |
| Quote-all CSV, 100,000 × 20 | 2,075.45 / 1,900.71 → 1,610.56 / 1,427.86 | 76.59 / 75.84 → 73.76 / 74.98 | -22.40% / -24.88% | -3.70% / -1.13% |
| Typed XLSX, 100,000 × 6 | 3,912.69 / 3,141.35 → 2,694.88 / 2,796.03 | 76.26 / 77.88 → 78.36 / 78.34 | -31.12% / -10.99% | +2.75% / +0.60% |

The completed wide Helvetica case contradicts the time criterion by 19.22% mean
and 12.08% median. The narrow Helvetica case contradicts the public-baseline RSS
criterion by 1.06% / 1.60%; its RSS is also 1.43% / 1.17% above M8. Wide
Helvetica RSS remains 17.86% / 13.66% below M8. Both quote-all CSV time
statistics improve, but the narrow quote-all RSS and XLSX RSS losses are retained.
These observations use all five measured samples; completing Carlito cannot
remove the already observed criterion failures.

Every completed Helvetica/CSV case retains all six exact public/candidate
output-hash matches, including its validated warmup. That does not close the
incomplete embedded-font byte and structure acceptance.

Canonical run identities:

| Stage | Run identity | Independently validated warmups / measured |
| --- | --- | ---: |
| Helvetica PDF | `20261024-132736-6c68578f` | 8 / 40 |
| CSV/XLSX | `20261024-140801-68027e6f` | 8 / 40 |
| Embedded Carlito PDF, incomplete | `20261024-141210-99b3346d` | 0 / 5 |

All raw measurement samples, validation receipts, retained partial producer
outputs and earlier M9/M10 losses remain separate. No sample is removed,
replaced, pooled across candidates or
replayed to obtain a favorable outcome. Background activity, launch variation,
CPU measurements and the supervisor timeout do not identify a new causal
performance component. The unchanged criteria remain unmet; this entry provides
no acceptance result for a later candidate and no universal speed, memory or
performance-no-regression claim.

## Page-command and complete-fragment lifetime

Measured 25 October 2026 UTC. Source
`900db5c4b97c87abe92901869ebdf16dbde16e73`, runtime SHA-256
`bfeee0cd16d32d445d0e6504a9d76106da20e2ed7a2f5cbd5a8f058c39592268`.

The third M10 candidate removes three source-proven common costs. Page
completion transfers its joined content directly to the existing stream writer
and releases the redundant command list before any codec or sink wait. Tagged
ordinary rows allocate row-header association lists only when an actual row
header is reached, while logical cell allocation and identifier order stay the
same. Complete ordinary first fragments reuse their captured physical cells and
line metrics; partial lines, spans, explicit heights and continuation artifacts
retain their existing paths. The minimum repeated-identifier scan is reused
within the same row operation.

These paths also exist in the controls. Removing their allocations does not
identify the cause of a particular timing or RSS difference. The measured
runtime and its complete inputs retain their original bytes and hashes; later
document or source edits do not reinterpret this cohort.

### Third M10 cohort: complete, criteria unmet

All **120 outputs** pass independent validation: **20 warmups and 100 measured
exports**, with one warmup and five measured samples per selected case/runtime.
Every changed PDF/CSV output remains byte-identical to public 0.1.2. The same
normal-host, rotated, no-outlier-removal policy and all five M10 criteria apply.
Helvetica and embedded Carlito remain separate profiles. Values below are
**mean / median**, with arrows showing **public 0.1.2 → third candidate**.
Times are milliseconds and RSS is MiB; changes use unrounded measurements.

| Workload | Export ms, public → candidate | Peak RSS MiB, public → candidate | Time change | RSS change |
| --- | ---: | ---: | ---: | ---: |
| Helvetica tagged PDF, 10,000 × 6 | 550.22 / 406.99 → 566.96 / 468.74 | 132.36 / 132.30 → 109.88 / 110.66 | +3.04% / +15.17% | -16.99% / -16.36% |
| Helvetica tagged PDF, 10,000 × 20 | 2,621.78 / 2,362.61 → 2,392.79 / 2,028.76 | 235.41 / 236.28 → 231.22 / 232.39 | -8.73% / -14.13% | -1.78% / -1.65% |
| Carlito tagged PDF, 10,000 × 6 | 541.38 / 362.89 → 876.34 / 1,052.36 | 134.64 / 134.80 → 116.46 / 117.27 | +61.87% / +189.99% | -13.50% / -13.01% |
| Carlito tagged PDF, 10,000 × 20 | 2,295.76 / 2,334.97 → 2,203.93 / 2,570.74 | 233.94 / 230.66 → 228.15 / 230.73 | -4.00% / +10.10% | -2.47% / +0.03% |
| Minimal CSV, 100,000 × 6 | 183.93 / 165.08 → 194.74 / 159.17 | 73.95 / 74.17 → 74.26 / 74.41 | +5.88% / -3.58% | +0.42% / +0.32% |
| Quote-all CSV, 100,000 × 6 | 229.70 / 251.86 → 184.00 / 148.05 | 74.83 / 74.70 → 74.48 / 74.34 | -19.89% / -41.22% | -0.47% / -0.48% |
| Quote-all CSV, 100,000 × 20 | 508.18 / 470.39 → 424.55 / 367.47 | 81.76 / 77.17 → 75.55 / 75.30 | -16.46% / -21.88% | -7.59% / -2.43% |
| Typed XLSX, 100,000 × 6 | 899.84 / 872.28 → 824.58 / 776.59 | 80.37 / 79.56 → 80.05 / 79.50 | -8.36% / -10.97% | -0.40% / -0.08% |

The time criterion remains **unmet**. Narrow Helvetica time is 3.04% / 15.17%
higher; narrow Carlito is 61.87% / 189.99% higher. Wide Carlito's mean improves,
but its median is 10.10% higher. Both quote-all CSV cases improve by mean and
median. Minimal CSV remains a guard observation: its mean is **5.88% higher**,
its median is 3.58% lower, and RSS is 0.42% / 0.32% higher. These losses are
retained rather than replaced with the favorable statistics.

The public-baseline RSS criterion also remains **unmet**. Wide Carlito's median
increases from 241,860,608 to 241,942,528 bytes: **+0.03387%**, or 81,920 bytes.
This small increase still fails the unchanged no-increase condition. Its mean
is 2.47% lower; both statistics of the other PDF cases are lower.

Both original M8 Helvetica memory criteria pass: narrow RSS mean/median is
18.52% / 17.83% lower, and wide is 21.86% / 15.66% lower. Same-cohort wide time
still has a 0.06% / 7.51% increase against M8, and a 12.39% mean reduction with
a **44.10% median increase against M9**. Those timing observations remain
separate from the M8 memory criterion. The XLSX guard has 8.36% / 10.97% lower
time and 0.40% / 0.08% lower RSS against public 0.1.2; its typed-value and schema
readbacks pass. Timestamp-dependent XLSX ZIP identity is not a byte-equality
requirement. Earlier XLSX and M9 losses stay bound to their earlier sources.

The unchanged criteria therefore resolve as follows:

| Criterion | Result |
| --- | --- |
| Independently validate all 120 actual outputs, including warmups | Pass |
| Preserve exact public-0.1.2 bytes for changed PDF/CSV cases | Pass |
| Lower selected PDF/quote-all CSV time by both mean and median | **Unmet** |
| PDF peak-RSS means and medians no higher than public 0.1.2 | **Unmet** |
| Preserve original Helvetica RSS means and medians against M8 | Pass |

Canonical run identities:

| Stage | Run identity | Independently validated warmups / measured |
| --- | --- | ---: |
| Helvetica PDF | `20261025-075623-3b1c0ffe` | 8 / 40 |
| CSV/XLSX | `20261025-081736-5dc4a5ae` | 8 / 40 |
| Embedded Carlito PDF | `20261025-081905-544038b0` | 4 / 20 |

All raw samples, output hashes and canonical summaries/comparisons remain
retained, including every warmup and loss. The first failed cohort, interrupted
second cohort and historical M9 comparisons are not pooled or substituted.
This normal-host five-sample result gives no universal speed, memory or
performance-no-regression guarantee.

### CPU diagnosis and separate string-accumulator experiment

Four CPU-profiled exports cover narrow Helvetica and Carlito under public 0.1.2
and the measured third runtime. All four preserve their corresponding cohort
bytes and pass qpdf. Whole-process samples include imports, font setup, layout
and output; they do not isolate the export timer or establish the cause of the
failed cohort statistics. Their perturbed time and memory observations are
not acceptance measurements and do not waive either unmet criterion.

The page-local string experiment is measured separately below. Its result does
not reinterpret the third cohort or waive either of that cohort's unmet criteria.

## Page-local string accumulation

Measured 25 October 2026 UTC. Source
`49c3f24dca8f5c4c653fcd28bc634d54a7a4f611`, runtime SHA-256
`299913df8b2bfd8367299304acac6b91e4077053fb0d1869418afe5c8ec3bd6f`.

The fourth M10 candidate appends drawing commands to one ordered page-local
string instead of retaining a command list and joining it at page completion.
It adds no public API, dependency or font cache. Page-byte accounting, stream
delivery, semantic ownership and resource limits retain their existing
contracts. Codec, arena, member and delivery-buffer policies are unchanged.
The measured runtime and all 19 measured inputs are retained with their original
hashes; subsequent source or documentation edits do not reinterpret this cohort.

### Fourth M10 cohort: complete, time criterion unmet

All **120 outputs** pass independent validation: **20 warmups and 100 measured
exports**, with one warmup and five measured samples per selected case/runtime.
Every changed PDF/CSV output is byte-identical to public 0.1.2. Public 0.1.2,
M8 and M9 use the same frozen compiled controls recorded above. The normal
shared-host, rotated-order policy and all five M10 criteria are unchanged;
no outliers are removed or samples pooled across candidates. Values below are
**mean / median**, with arrows showing **public 0.1.2 → fourth candidate**.
Times are milliseconds and RSS is MiB; changes use unrounded measurements.

| Workload | Export ms, public → candidate | Peak RSS MiB, public → candidate | Time change | RSS change |
| --- | ---: | ---: | ---: | ---: |
| Helvetica tagged PDF, 10,000 × 6 | 982.87 / 992.06 → 884.71 / 927.27 | 131.81 / 132.23 → 111.33 / 110.95 | -9.99% / -6.53% | -15.53% / -16.09% |
| Helvetica tagged PDF, 10,000 × 20 | 2,083.73 / 1,973.73 → 3,407.71 / 3,376.80 | 239.91 / 238.95 → 228.67 / 229.11 | +63.54% / +71.09% | -4.69% / -4.12% |
| Carlito tagged PDF, 10,000 × 6 | 1,374.16 / 1,343.82 → 1,264.68 / 1,333.76 | 135.87 / 135.97 → 113.45 / 112.44 | -7.97% / -0.75% | -16.50% / -17.31% |
| Carlito tagged PDF, 10,000 × 20 | 3,752.72 / 4,033.08 → 4,012.76 / 4,100.15 | 230.41 / 231.02 → 230.39 / 227.81 | +6.93% / +1.66% | -0.01% / -1.39% |
| Minimal CSV, 100,000 × 6 | 582.31 / 610.01 → 572.59 / 544.89 | 74.03 / 74.11 → 74.46 / 74.61 | -1.67% / -10.68% | +0.58% / +0.67% |
| Quote-all CSV, 100,000 × 6 | 540.02 / 564.22 → 532.26 / 502.45 | 74.66 / 74.67 → 74.64 / 74.73 | -1.44% / -10.95% | -0.02% / +0.08% |
| Quote-all CSV, 100,000 × 20 | 1,695.90 / 1,873.66 → 1,437.37 / 1,246.67 | 76.91 / 75.31 → 75.09 / 74.94 | -15.24% / -33.46% | -2.36% / -0.50% |
| Typed XLSX, 100,000 × 6 | 2,908.11 / 2,805.76 → 2,800.10 / 2,874.30 | 78.75 / 78.67 → 78.88 / 78.47 | -3.71% / +2.44% | +0.16% / -0.26% |

Only the selected time criterion remains unmet in this cohort. Narrow Helvetica
time is 9.99% / 6.53% lower and narrow Carlito is 7.97% / 0.75% lower than
public 0.1.2. Both quote-all CSV time statistics also improve. Wide Helvetica,
however, is **63.54% / 71.09% slower**, and wide Carlito is **6.93% / 1.66%
slower**. These increases are retained without a waiver.

All four PDF RSS means and medians are no higher than public 0.1.2. Wide
Carlito's mean reduction is only 0.00678%, or 16,384 bytes; this is not evidence
of a material memory improvement. Both original M8 Helvetica memory criteria
also pass: narrow RSS mean/median is 16.65% / 16.89% lower and wide is
11.26% / 7.85% lower. Same-cohort wide Helvetica time is 27.74% / 27.06% lower
than M8 but **18.56% / 23.44% higher than M9**. These different baselines do not
replace the failed public-baseline time criterion.

Guard losses remain visible. Minimal-CSV RSS increases 0.58% / 0.67%; narrow
quote-all CSV has a 0.08% higher RSS median. Typed XLSX has a 3.71% lower time
mean but a **2.44% higher median**, with RSS mean 0.16% higher and median 0.26%
lower. Typed-value and schema readbacks pass; timestamp-dependent XLSX ZIP
identity is not a byte-equality requirement. Earlier minimal-CSV, XLSX and PDF
losses remain bound to their earlier cohorts.

| Criterion | Result |
| --- | --- |
| Independently validate all 120 actual outputs, including warmups | Pass |
| Preserve exact public-0.1.2 bytes for changed PDF/CSV cases | Pass |
| Lower selected PDF/quote-all CSV time by both mean and median | **Unmet** |
| PDF peak-RSS means and medians no higher than public 0.1.2 | Pass |
| Preserve original Helvetica RSS means and medians against M8 | Pass |

Canonical run identities:

| Stage | Run identity | Independently validated warmups / measured |
| --- | --- | ---: |
| Helvetica PDF | `20261025-085228-00492158` | 8 / 40 |
| CSV/XLSX | `20261025-091820-e592eaac` | 8 / 40 |
| Embedded Carlito PDF | `20261025-092132-685bc621` | 4 / 20 |

All raw samples, warmups, output hashes and canonical summaries/comparisons are
retained, including the failed and interrupted prior cohorts. Normal shared-host
activity remains part of this observation; no quiet-host, priority, affinity or
forced-GC treatment is inferred. This five-sample result provides no universal
speed, memory or performance-no-regression guarantee.

### Wide-PDF diagnosis and separate row-string experiment

Four additional CPU-profiled exports cover wide Helvetica and Carlito under
public 0.1.2 and the measured fourth runtime. All four preserve their cohort
bytes and pass qpdf. Samples move between page completion and encoding work;
whole-process self-sample counts include imports, setup, layout and output.
They neither isolate a causal regression nor substitute for the failed time
criterion. Instrumented timing and RSS are not acceptance measurements.

The physical-row command-parts experiment is measured separately below. Its
result does not reinterpret the fourth cohort or waive that cohort's time loss.

## Physical-row command parts

Measured 25 October 2026 UTC. Source
`5b2bb0f76aa6a41bec99f2e505ea19953245dc80`, runtime SHA-256
`f06a55a22393ffccd42d90824246b4399c5f034e9af0261e2e4b5e4ad21c0094`.

The fifth M10 candidate accumulates drawing commands in a string tail and seals
that tail into ordered parts at existing physical-row boundaries. Page completion
joins the parts and any trailing commands, then releases both before awaiting
delivery. Whole-page byte accounting and command order are unchanged; sealing
a row adds no intermediate sink write. The font mapping, byte codec, public API,
dependencies, cancellation and backpressure contracts are unchanged.

### Fifth M10 cohort: complete, time and public-RSS criteria unmet

All **120 actual outputs** pass independent validation: **20 warmups and 100
measured exports**, with one warmup and five measured samples per selected
case/runtime. Every changed PDF/CSV output is byte-identical to public 0.1.2.
The same frozen public 0.1.2, M8 and M9 controls, rotated-order policy and five
unchanged M10 criteria apply. Values below are **mean / median**; arrows show
**public 0.1.2 → fifth candidate**. Times are milliseconds and RSS is MiB;
percentage changes use unrounded measurements.

| Workload | Export ms, public → candidate | Peak RSS MiB, public → candidate | Time change | RSS change |
| --- | ---: | ---: | ---: | ---: |
| Helvetica tagged PDF, 10,000 × 6 | 1,607.96 / 1,469.74 → 1,444.15 / 1,447.09 | 129.82 / 131.02 → 106.42 / 107.89 | -10.19% / -1.54% | -18.02% / -17.65% |
| Helvetica tagged PDF, 10,000 × 20 | 3,758.26 / 3,977.68 → 3,217.59 / 3,242.68 | 233.96 / 234.69 → 221.12 / 219.02 | -14.39% / -18.48% | -5.49% / -6.68% |
| Carlito tagged PDF, 10,000 × 6 | 3,072.41 / 1,905.82 → 2,010.37 / 1,615.87 | 132.21 / 132.73 → 114.48 / 117.53 | -34.57% / -15.21% | -13.41% / -11.45% |
| Carlito tagged PDF, 10,000 × 20 | 4,122.48 / 3,448.13 → 4,021.18 / 4,012.89 | 236.21 / 237.92 → 234.86 / 238.47 | -2.46% / +16.38% | -0.57% / +0.23% |
| Minimal CSV, 100,000 × 6 | 439.82 / 374.63 → 635.82 / 525.96 | 74.28 / 74.22 → 73.02 / 74.55 | +44.56% / +40.40% | -1.71% / +0.44% |
| Quote-all CSV, 100,000 × 6 | 837.95 / 661.64 → 701.71 / 686.10 | 74.55 / 74.61 → 74.71 / 74.66 | -16.26% / +3.70% | +0.22% / +0.06% |
| Quote-all CSV, 100,000 × 20 | 2,780.67 / 2,676.87 → 1,866.04 / 1,754.31 | 72.62 / 73.66 → 74.14 / 73.73 | -32.89% / -34.46% | +2.10% / +0.11% |
| Typed XLSX, 100,000 × 6 | 2,526.68 / 2,575.11 → 3,096.06 / 3,305.39 | 76.60 / 76.89 → 78.06 / 78.17 | +22.53% / +28.36% | +1.91% / +1.67% |

Both Helvetica PDF time statistics improve: narrow is 10.19% / 1.54% lower and
wide is 14.39% / 18.48% lower than public 0.1.2. Narrow Carlito is 34.57% /
15.21% lower, and wide quote-all CSV is 32.89% / 34.46% lower. These improvements
do not satisfy a criterion requiring every selected case to pass.

Wide Carlito has a 2.46% lower time mean but a **16.38% higher median**. Its
RSS mean is 0.57% lower, while its median rises from 249,479,168 to 250,052,608
bytes: **+0.22985%**, or 573,440 bytes. That increase fails the unchanged
public-baseline PDF RSS criterion. Narrow quote-all CSV has a 16.26% lower
time mean but a **3.70% higher median**, so it also fails the time criterion.
Its CPU mean/median is 8.13% / 9.17% lower; that separate metric does not waive
the export-time median loss.

The original M8 Helvetica RSS criterion passes: narrow mean/median is 18.38% /
16.14% lower, and wide is 14.21% / 14.58% lower. All PDF/CSV byte and independent
structure/value checks pass. These results do not erase the fifth cohort's
public-RSS and time failures or any earlier failed memory measurement.

Guard losses remain material. Minimal-CSV time is **44.56% / 40.40% higher**,
with CPU 7.69% / 14.83% higher; its RSS mean is 1.71% lower but median 0.44%
higher. Typed XLSX time is **22.53% / 28.36% higher**, CPU is 2.67% / 1.00%
higher and RSS is 1.91% / 1.67% higher. Wide quote-all CSV's time improves while
RSS rises 2.10% / 0.11%; narrow quote-all RSS rises 0.22% / 0.06%. XLSX typed
values and schema pass; timestamp-dependent ZIP identity remains outside the
PDF/CSV byte-equality requirement.

| Criterion | Result |
| --- | --- |
| Independently validate all 120 actual outputs, including warmups | Pass |
| Preserve exact public-0.1.2 bytes for changed PDF/CSV cases | Pass |
| Lower selected PDF/quote-all CSV time by both mean and median | **Unmet** |
| PDF peak-RSS means and medians no higher than public 0.1.2 | **Unmet** |
| Preserve original Helvetica RSS means and medians against M8 | Pass |

Canonical run identities:

| Stage | Run identity | Independently validated warmups / measured |
| --- | --- | ---: |
| Helvetica PDF | `20261025-094925-12108486` | 8 / 40 |
| CSV/XLSX | `20261025-102030-6bf39d19` | 8 / 40 |
| Embedded Carlito PDF | `20261025-102352-12efe864` | 4 / 20 |

The complete frozen inputs, all raw samples, warmups, output hashes and canonical
summaries/comparisons are retained. No sample is discarded, replaced or pooled
with an earlier candidate. This is a normal shared-host observation, without
quiet-host treatment or outlier selection. Background activity is not a causal
explanation or a waiver for a failed criterion. No universal speed, memory or
performance-no-regression claim follows from this five-sample cohort.

### CID and CSV diagnosis before the sixth cohort

Two further CPU-profiled exports preserve their corresponding bytes and pass
qpdf. They are diagnostics, not acceptance timings or proof of a causal
regression. The combined CID-text/CSV-dispatch candidate is measured separately
below. The diagnostic outputs do not isolate the contribution of either change
or reinterpret the fifth cohort's failed time and public-RSS criteria.

## Mapping-local CID text and primitive CSV dispatch

Measured 25 October 2026 UTC. Source
`94209992006b2b1cbc799a3a310eb1f774748002`, runtime SHA-256
`ec1c2718f79d5e7bd635e7574b0684cfeb7e772e2b537235d58d5207da7a4608`.

The sixth M10 candidate reuses a four-digit CID hexadecimal string on each
existing font-mapping entry at both CID registration owners. It introduces no
new map or cross-document state. CSV captures the primitive kind once and
handles strings, numbers and booleans before native Date recognition, while
retaining Date validation, error, formula-protection and quoting contracts.
The fifth candidate's physical-row command ownership is retained. There is no
public API or dependency change. Mapping-local text adds bounded per-glyph
state, so memory remains measured rather than assumed to improve.

### Sixth M10 cohort: complete, time criterion unmet

All **120 actual outputs** pass independent validation: **20 warmups and 100
measured exports**, with one warmup and five measured samples per selected
case/runtime. Every changed PDF/CSV output is byte-identical to public 0.1.2.
All 21 frozen inputs remain unchanged. The same public 0.1.2, M8 and M9 controls,
rotated order, normal shared-host conditions and five unchanged M10 criteria
apply. Values below are **mean / median**; arrows show **public 0.1.2 → sixth
candidate**. Times are milliseconds and RSS is MiB; percentage changes use
unrounded measurements.

| Workload | Export ms, public → candidate | Peak RSS MiB, public → candidate | Time change | RSS change |
| --- | ---: | ---: | ---: | ---: |
| Helvetica tagged PDF, 10,000 × 6 | 536.30 / 448.58 → 447.12 / 348.44 | 131.59 / 132.00 → 110.17 / 109.84 | -16.63% / -22.32% | -16.28% / -16.79% |
| Helvetica tagged PDF, 10,000 × 20 | 1,633.29 / 1,649.75 → 1,555.13 / 1,227.74 | 238.94 / 240.84 → 231.33 / 232.98 | -4.79% / -25.58% | -3.18% / -3.26% |
| Carlito tagged PDF, 10,000 × 6 | 568.95 / 361.35 → 413.58 / 362.90 | 133.95 / 133.11 → 108.22 / 107.47 | -27.31% / +0.43% | -19.21% / -19.26% |
| Carlito tagged PDF, 10,000 × 20 | 1,612.12 / 1,340.87 → 1,875.51 / 1,644.47 | 237.26 / 239.14 → 236.38 / 236.00 | +16.34% / +22.64% | -0.37% / -1.31% |
| Minimal CSV, 100,000 × 6 | 174.43 / 145.46 → 169.86 / 140.63 | 73.89 / 73.77 → 73.83 / 73.75 | -2.62% / -3.32% | -0.08% / -0.02% |
| Quote-all CSV, 100,000 × 6 | 195.84 / 160.95 → 219.97 / 191.89 | 74.26 / 74.31 → 74.85 / 74.95 | +12.32% / +19.23% | +0.80% / +0.86% |
| Quote-all CSV, 100,000 × 20 | 431.00 / 426.31 → 418.96 / 422.12 | 86.80 / 89.12 → 75.15 / 75.41 | -2.79% / -0.98% | -13.42% / -15.39% |
| Typed XLSX, 100,000 × 6 | 816.12 / 695.66 → 1,139.27 / 703.45 | 80.38 / 80.45 → 79.16 / 79.30 | +39.60% / +1.12% | -1.51% / -1.44% |

Only the selected time criterion fails in this cohort. Helvetica narrow time
is 16.63% / 22.32% lower and wide is 4.79% / 25.58% lower than public 0.1.2.
Narrow Carlito's mean is 27.31% lower, but its median rises from 361.349666 to
362.902125 ms: **+0.42963%**. Wide Carlito is **16.34% / 22.64% slower**.
Neither the small narrow-median increase nor the larger wide losses are waived.

Narrow quote-all CSV is **12.32% / 19.23% slower**, with CPU 5.45% / 22.56%
higher and RSS 0.80% / 0.86% higher. Minimal CSV improves by 2.62% / 3.32% time
and 0.71% / 2.80% CPU. Wide quote-all CSV improves by 2.79% / 0.98% time,
4.33% / 3.12% CPU and 13.42% / 15.39% RSS. These different outcomes do not
isolate the CSV dispatch change as either a causal gain or a regression cause.

All four PDF peak-RSS means and medians are no higher than public 0.1.2.
Both original M8 Helvetica memory criteria also pass: narrow RSS mean/median
is 17.82% / 18.05% lower and wide is 15.33% / 14.38% lower. Exact changed
PDF/CSV bytes and independent value/structure checks pass, without replacing
any earlier failed memory result.

The typed XLSX guard retains a **39.60% higher time mean and 1.12% higher
median**. Its CPU mean is 12.33% higher while its median is 0.04149% lower;
RSS is 1.51% / 1.44% lower. Typed-value and schema readbacks pass, while
timestamp-dependent ZIP identity remains outside the PDF/CSV byte-equality
requirement. Earlier XLSX and CSV losses remain visible in their own cohorts.

| Criterion | Result |
| --- | --- |
| Independently validate all 120 actual outputs, including warmups | Pass |
| Preserve exact public-0.1.2 bytes for changed PDF/CSV cases | Pass |
| Lower selected PDF/quote-all CSV time by both mean and median | **Unmet** |
| PDF peak-RSS means and medians no higher than public 0.1.2 | Pass |
| Preserve original Helvetica RSS means and medians against M8 | Pass |

Canonical run identities:

| Stage | Run identity | Independently validated warmups / measured |
| --- | --- | ---: |
| Helvetica PDF | `20261025-110618-5d24179d` | 8 / 40 |
| CSV/XLSX | `20261025-112535-35862516` | 8 / 40 |
| Embedded Carlito PDF | `20261025-112722-48f11425` | 4 / 20 |

The frozen inputs, every measurement and warmup sample, output hashes,
validation receipts and canonical summaries remain retained. The runner removes
successfully validated bulk artifacts; decisive diagnostic outputs are kept.
No sample is discarded, pooled across candidates or replayed to obtain a
favorable result. No background-activity explanation or CPU
measurement waives failed export-time criteria. These normal shared-host,
five-sample observations provide no universal speed, memory or
performance-no-regression guarantee and qualify no later source change.

## Numeric CSV and font selection: seventh M10 cohort

Measured 25 October 2026 UTC. Source
`3d1fcc99f4d4b2995211a89858c6eca172d3069d`, runtime SHA-256
`1d2c28dd6f35ed4e560888aa68ffbf61107d8ac9820b10f8252ef602e412bed9`.

The seventh candidate includes primitive-number CSV completion before common
string escaping. It preserves the existing finite-number conversion and empty
non-finite result, wrapping either result for quote-all and leaving numbers
unquoted in the other modes. String-returning formatters, null replacements,
Dates, formula protection, source ownership and stream behavior retain their
existing paths. Font selection also uses a private numeric key for the four bold/italic
combinations, avoiding repeated boolean-to-string conversion. Explicit-font
registration, shared font programs and first-use resource order are unchanged.
The complete runtime is measured together; this cohort does not isolate the
contribution of an individual source change.

All **120 actual outputs** pass independent validation: **20 warmups and 100
measured exports**, with one warmup and five measured samples per selected
case/runtime. Every changed PDF/CSV output is byte-identical to public 0.1.2.
The same frozen controls, rotated order, normal shared-host conditions and
five unchanged M10 criteria apply. Every numeric pair below is **mean / median**;
arrows show **public 0.1.2 → seventh candidate**. CPU is process CPU time and
RSS is lifetime process peak memory. Percentage changes use unrounded values.

### Export time

| Workload | Export time, public → candidate (ms) | Change |
| --- | ---: | ---: |
| Helvetica tagged PDF, 10,000 × 6 | 1,150.39 / 1,163.55 → 1,162.99 / 1,085.91 | +1.09% / -6.67% |
| Helvetica tagged PDF, 10,000 × 20 | 3,100.16 / 3,126.22 → 2,826.87 / 2,706.31 | -8.82% / -13.43% |
| Carlito tagged PDF, 10,000 × 6 | 1,253.01 / 1,043.52 → 926.79 / 918.98 | -26.03% / -11.93% |
| Carlito tagged PDF, 10,000 × 20 | 2,333.84 / 2,307.36 → 1,866.77 / 1,851.13 | -20.01% / -19.77% |
| Minimal CSV, 100,000 × 6 | 384.53 / 380.08 → 382.69 / 381.13 | -0.48% / +0.28% |
| Quote-all CSV, 100,000 × 6 | 395.75 / 412.31 → 355.06 / 356.47 | -10.28% / -13.54% |
| Quote-all CSV, 100,000 × 20 | 1,187.06 / 1,183.98 → 964.28 / 935.36 | -18.77% / -21.00% |
| Typed XLSX, 100,000 × 6 | 1,724.27 / 1,722.39 → 1,771.73 / 1,768.08 | +2.75% / +2.65% |

### Process CPU time

| Workload | CPU time, public → candidate (ms) | Change |
| --- | ---: | ---: |
| Helvetica tagged PDF, 10,000 × 6 | 1,064.89 / 1,068.99 → 973.53 / 1,007.66 | -8.58% / -5.74% |
| Helvetica tagged PDF, 10,000 × 20 | 3,078.99 / 3,114.48 → 2,581.65 / 2,679.66 | -16.15% / -13.96% |
| Carlito tagged PDF, 10,000 × 6 | 1,128.64 / 1,172.09 → 936.79 / 956.99 | -17.00% / -18.35% |
| Carlito tagged PDF, 10,000 × 20 | 2,424.76 / 2,460.28 → 2,064.58 / 2,140.74 | -14.85% / -12.99% |
| Minimal CSV, 100,000 × 6 | 332.62 / 349.57 → 351.56 / 352.17 | +5.70% / +0.74% |
| Quote-all CSV, 100,000 × 6 | 398.52 / 408.08 → 366.67 / 364.49 | -7.99% / -10.68% |
| Quote-all CSV, 100,000 × 20 | 1,003.02 / 1,030.54 → 851.22 / 840.01 | -15.13% / -18.49% |
| Typed XLSX, 100,000 × 6 | 1,527.04 / 1,512.23 → 1,575.93 / 1,564.89 | +3.20% / +3.48% |

### Peak RSS

| Workload | Peak RSS, public → candidate (MiB) | Change |
| --- | ---: | ---: |
| Helvetica tagged PDF, 10,000 × 6 | 132.35 / 132.34 → 106.83 / 107.48 | -19.28% / -18.78% |
| Helvetica tagged PDF, 10,000 × 20 | 231.12 / 230.58 → 226.32 / 229.23 | -2.07% / -0.58% |
| Carlito tagged PDF, 10,000 × 6 | 132.77 / 133.70 → 108.65 / 108.50 | -18.17% / -18.85% |
| Carlito tagged PDF, 10,000 × 20 | 236.97 / 236.55 → 229.32 / 231.77 | -3.23% / -2.02% |
| Minimal CSV, 100,000 × 6 | 73.94 / 74.02 → 74.35 / 74.38 | +0.56% / +0.49% |
| Quote-all CSV, 100,000 × 6 | 74.59 / 74.47 → 74.49 / 74.30 | -0.13% / -0.23% |
| Quote-all CSV, 100,000 × 20 | 76.41 / 76.06 → 75.03 / 75.20 | -1.80% / -1.13% |
| Typed XLSX, 100,000 × 6 | 79.03 / 79.30 → 78.88 / 78.84 | -0.19% / -0.57% |

The sole selected-speed failure is narrow Helvetica's export-time mean:
1,150.3924248 → 1,162.986525 ms, **+1.09477%**. Its median improves 6.67%,
CPU falls 8.58% / 5.74% and RSS falls 19.28% / 18.78%. Those separate statistics
do not waive the higher mean. Both statistics improve for each of the other
five selected speed cases: wide Helvetica, narrow/wide Carlito and narrow/wide
quote-all CSV.

All four PDF peak-RSS means and medians are no higher than public 0.1.2.
Both original M8 Helvetica memory criteria also pass: narrow RSS mean/median
is 20.24% / 20.13% lower, and wide is 10.47% / 10.47% lower. The same-cohort
M9 comparison retains a narrow-PDF median increase of **8.53%**, despite a
5.39% lower mean. Each baseline keeps its own meaning; neither memory gains
nor a different comparison replace the failed public-baseline time mean.

Guard losses remain explicit. Minimal CSV has a 0.48% lower time mean but a
**0.28% higher median**, CPU 5.70% / 0.74% higher and RSS 0.56% / 0.49% higher.
Typed XLSX time is **2.75% / 2.65% higher** and CPU is 3.20% / 3.48% higher,
while RSS is 0.19% / 0.57% lower. Its typed-value and schema checks pass;
timestamp-dependent XLSX ZIP identity remains outside the PDF/CSV byte-equality
criterion. These guard observations and all earlier mixed results are retained
without a universal performance claim or a causal explanation inferred from
the cohort alone.

| Criterion | Result |
| --- | --- |
| Independently validate all 120 actual outputs, including warmups | Pass |
| Preserve exact public-0.1.2 bytes for changed PDF/CSV cases | Pass |
| Lower selected PDF/quote-all CSV time by both mean and median | **Unmet** |
| PDF peak-RSS means and medians no higher than public 0.1.2 | Pass |
| Preserve original Helvetica RSS means and medians against M8 | Pass |

Canonical run identities:

| Stage | Run identity | Independently validated warmups / measured |
| --- | --- | ---: |
| Helvetica PDF | `20261025-115816-cc6e43f5` | 8 / 40 |
| CSV/XLSX | `20261025-122831-8b9ff656` | 8 / 40 |
| Embedded Carlito PDF | `20261025-123059-94775740` | 4 / 20 |

Raw measurements, all warmup measurement records, output hashes, independent
validation receipts and canonical summaries remain retained. The canonical
runner automatically deletes successfully validated bulk output files;
retention does not mean that every PDF, CSV or XLSX binary is still present.
Decisive diagnostic artifacts are retained separately. No sample is discarded,
replaced, pooled across candidates or rerun to select favorable statistics.
The time criterion remains unmet, and these five-sample observations provide
no universal speed, memory or performance-no-regression guarantee and no
acceptance for a later source change.

## Reused RGB command text: eighth M10 cohort

Measured 25 October 2026 UTC. Source
`959ca451132fb20b1f3302f06baa303bf2b0a0cb`, runtime SHA-256
`720aab4e0cfb0f9397b1f7f3e10c53c1d890ff27c8a677b5b142e96ba2a4b4d2`.

The eighth candidate reuses one validated RGB string within each ordinary or
shaped PDF text command. It retains the preceding font-selection and CSV paths.
The complete runtime is measured together; these results do not isolate an RGB
serialization gain or establish the cause of a slower sample.

The fixed cohort completes with exit code zero in **2,380.517 seconds**. All
**120 actual outputs** pass independent validation: **20 warmups and 100 measured
exports**, with one warmup and five measured samples per selected case/runtime.
Every changed PDF/CSV output is byte-identical to public 0.1.2. The same frozen
controls, rotated order, normal shared-host conditions and five unchanged M10
criteria apply. Every pair below is **mean / median**; arrows show **public
0.1.2 → eighth candidate**. Export time is the recorded `EngineMs`, rather than
the outer process duration. CPU is aggregate process `NodeCpuMs`, including
native threads. RSS is lifetime process peak memory; its table keeps the exact
recorded bytes. Percentage changes use unrounded values.

### Export time

| Workload | Export time, public → candidate (ms) |
| --- | ---: |
| Helvetica tagged PDF, 10,000 × 6 | 657.20 / 740.49 → 778.60 / 742.09 |
| Helvetica tagged PDF, 10,000 × 20 | 3,291.41 / 2,962.61 → 2,481.56 / 2,492.42 |
| Carlito tagged PDF, 10,000 × 6 | 960.49 / 1,069.65 → 827.15 / 807.02 |
| Carlito tagged PDF, 10,000 × 20 | 3,370.43 / 3,418.47 → 2,525.14 / 2,467.35 |
| Minimal CSV, 100,000 × 6 | 281.76 / 229.82 → 245.29 / 212.66 |
| Quote-all CSV, 100,000 × 6 | 217.05 / 203.09 → 184.82 / 187.45 |
| Quote-all CSV, 100,000 × 20 | 512.91 / 545.92 → 485.14 / 506.45 |
| Typed XLSX, 100,000 × 6 | 1,083.69 / 910.93 → 868.35 / 902.01 |

### Process CPU time

| Workload | CPU time, public → candidate (ms) |
| --- | ---: |
| Helvetica tagged PDF, 10,000 × 6 | 785.75 / 878.18 → 801.72 / 839.09 |
| Helvetica tagged PDF, 10,000 × 20 | 2,659.83 / 2,712.27 → 2,349.23 / 2,259.72 |
| Carlito tagged PDF, 10,000 × 6 | 1,022.23 / 1,134.58 → 848.20 / 934.56 |
| Carlito tagged PDF, 10,000 × 20 | 3,200.74 / 3,210.86 → 2,543.36 / 2,555.46 |
| Minimal CSV, 100,000 × 6 | 262.73 / 263.47 → 245.60 / 251.48 |
| Quote-all CSV, 100,000 × 6 | 237.50 / 241.58 → 218.98 / 226.06 |
| Quote-all CSV, 100,000 × 20 | 542.85 / 587.67 → 507.58 / 534.74 |
| Typed XLSX, 100,000 × 6 | 1,010.78 / 933.94 → 932.43 / 944.93 |

### Peak RSS

| Workload | Peak RSS, public → candidate (bytes) |
| --- | ---: |
| Helvetica tagged PDF, 10,000 × 6 | 138,543,104 / 138,297,344 → 113,777,049.6 / 115,310,592 |
| Helvetica tagged PDF, 10,000 × 20 | 249,652,838.4 / 248,840,192 → 243,489,177.6 / 243,351,552 |
| Carlito tagged PDF, 10,000 × 6 | 140,833,587.2 / 140,820,480 → 113,508,352 / 113,065,984 |
| Carlito tagged PDF, 10,000 × 20 | 241,093,836.8 / 241,041,408 → 237,174,784 / 236,732,416 |
| Minimal CSV, 100,000 × 6 | 77,771,571.2 / 77,725,696 → 77,879,705.6 / 77,824,000 |
| Quote-all CSV, 100,000 × 6 | 78,361,395.2 / 78,315,520 → 77,955,072 / 77,922,304 |
| Quote-all CSV, 100,000 × 20 | 85,698,150.4 / 81,461,248 → 79,151,104 / 79,151,104 |
| Typed XLSX, 100,000 × 6 | 83,663,257.6 / 83,148,800 → 83,895,910.4 / 83,542,016 |

Narrow Helvetica fails both selected time statistics: its mean rises from
657.1978002 to 778.6005 ms, **+18.47278%**, and its median rises from 740.494083
to 742.090667 ms, **+0.21561%**. CPU mean is 2.03% higher while its median is
4.45% lower. Neither CPU nor memory gains waive either export-time loss. Both
time statistics improve for the other five selected speed cases: wide
Helvetica, narrow/wide Carlito and narrow/wide quote-all CSV.

All four PDF peak-RSS means and medians are no higher than public 0.1.2.
The original M8 Helvetica memory criteria also pass: narrow RSS is 19.79% /
19.12% lower and wide is 11.15% / 12.59% lower. M9 remains a separate control:
narrow Helvetica time is **1.95% / 111.56% higher** and wide is **25.13% /
51.38% higher**. The narrow M8 time median is also 43.65% higher. These losses
are retained alongside the public-baseline gains and memory results.

The unchanged guards retain mixed results. Minimal CSV improves time and CPU,
but RSS rises **0.13904% / 0.12648%**. Typed XLSX time improves **19.87% /
0.98%**, while CPU median rises **1.17695%** and RSS rises **0.27808% /
0.47291%**. Its typed-value and schema readbacks pass; timestamp-dependent ZIP
identity remains outside the changed PDF/CSV byte-equality criterion. These
results do not replace earlier CSV or XLSX losses.

| Criterion | Result |
| --- | --- |
| Independently validate all 120 actual outputs, including warmups | Pass |
| Preserve exact public-0.1.2 bytes for changed PDF/CSV cases | Pass |
| Lower selected PDF/quote-all CSV time by both mean and median | **Unmet** |
| PDF peak-RSS means and medians no higher than public 0.1.2 | Pass |
| Preserve original Helvetica RSS means and medians against M8 | Pass |

Canonical run identities:

| Stage | Run identity | Independently validated warmups / measured |
| --- | --- | ---: |
| Helvetica PDF | `20261025-141215-a74112b8` | 8 / 40 |
| CSV/XLSX | `20261025-143724-a2cb9eed` | 8 / 40 |
| Embedded Carlito PDF | `20261025-143907-72370a0a` | 4 / 20 |

Frozen inputs, raw samples, warmup measurement records, output hashes,
independent validation receipts and canonical reports remain retained. The
canonical runner automatically deletes successfully validated bulk output
files, so retention does not mean every PDF, CSV or XLSX binary remains.
All seven earlier cohorts stay separate and unwaived. No sample is discarded,
pooled across candidates or replayed to select favorable statistics; background
activity does not explain away a failed criterion. These five-sample
shared-host observations establish no universal speed, memory or
performance-no-regression guarantee.

The later single-entry PDF command-prefix experiment has runtime SHA-256
`564a689cd48631bf69c9fc9c13167faabe7923fa5db38042c8d5a9b7f8a303c8`.
It is a different candidate from the measured RGB reuse above and has no timing
or memory acceptance from this cohort.
