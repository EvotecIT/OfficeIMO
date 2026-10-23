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

First qualify actual header associations, repeated/reordered/banded tables,
compressed and classic fallback, cancellation and limits, then independently
read narrow/wide outputs. Only then run one fixed baseline/current cohort.
Repeated decoded payload is proven; a memory benefit is still a hypothesis.
Do not repeat an unchanged candidate to obtain a favorable number, enlarge the
decoded arena or substitute file-size savings for the narrow-memory target.
