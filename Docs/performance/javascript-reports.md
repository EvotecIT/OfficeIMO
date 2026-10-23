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
not isolate a separate RSS gain for either lifetime change. Publication remains
held because narrow memory acceptance is unmet.

## Next comparison and stop rule

The next bounded experiment changes the member ceiling from 200 to 256 while
retaining the exact 32 KiB decoded-byte ceiling, arena size and sink chunks.
The current packing now reaches 200 members in every full batch; this cause
was previously obscured by page-induced partial batches. Read-only packing of
retained member bytes projects 315 narrow and 1,025 wide containers. Actual ID
widths can change, so those counts are a hypothesis, not acceptance. The
README/architecture member bound is revised explicitly; no public option,
provider or dependency is added. The two-byte xref member index must resolve
index 255 and repeated arena reuse correctly.

First check actual container counts, bounds, page content, logical order,
headers, ParentTree/IDTree, links, fallback, cancellation, output limits and
slow sinks. Only then run one fixed baseline/current narrow-and-wide cohort.
Do not repeat an unchanged candidate to obtain a favorable number. If this
capacity experiment still misses narrow memory acceptance, stop tuning the
member ceiling and retain the loss. Any next experiment needs a different
measured cause; do not silently enlarge the decoded arena or waive the target.
