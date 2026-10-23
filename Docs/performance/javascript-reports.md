# JavaScript analytical-report performance evidence

## Workload and measurement boundary

The current investigation targets peak process memory for tagged, streamed
analytical PDFs. It preserves the complete logical table, header associations,
identifiers, page content and navigation. Smaller files alone do not satisfy the
narrow-report memory target.

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

## Accepted 23 October 2026 checkpoint

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
