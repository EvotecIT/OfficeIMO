# iWork runtime comparison — 2026-10-04

The repeated synthetic matrix records lower managed allocation for Numbers and
Keynote, and faster large Pages and Keynote conversions. Several loading and
smaller conversion cases took longer. These are observations on one macOS host;
they do not establish portable elapsed-time or memory budgets.

The [comparison and assembly hashes](iwork-runtime-2026-10-04/comparison.json) and
all four raw sample files retain the measured results. The baseline owner source
is `683cf842502771c652b9d7b14cd8bf0738b4fca1`. Candidate assemblies were built from
`8a2f135e1f286bfed49a0428f38e4269048e46fd` with the staged Keynote patch identified
in that manifest, now committed as `44c69d1c18f415e2967173b70d2175de7e5115f2`.
The assemblies, input hashes and selected conversion policy
identify the measured work; the runner's checkout alone does not identify a
baseline assembly loaded from another directory.

## Work and measurement

The [opt-in workload](../../OfficeIMO.IWork.Benchmarks/README.md) measures
`LoadProject` and `ConvertSave` for each family at three sizes. Pages uses
100/1,000/10,000 paragraphs, Numbers 1,000/10,000/100,000 numeric cells and Keynote
10/100/1,000 slides. Input creation and full semantic readback stay outside timing.
Every saved DOCX, XLSX and PPTX is reopened and checked. Synthetic conversions
require complete editable reconstruction. All 360 measured samples succeeded,
with matching input hashes and verified content across variants.

Each matrix ran in a fresh PowerShell 7.5.4 process on .NET 9.0.10, loading the
net8.0 workload and owner assemblies built with SDK 10.0.112. Runs alternated
baseline, candidate, candidate, baseline. Each case had two warmups and five
measured iterations, rotated case order, collection before each iteration and no
outlier removal. The table takes the median of each variant's two run medians.
Negative changes mean fewer bytes or less elapsed time.

| Case | Operation | Baseline ms | Candidate ms | Elapsed change | Allocation change |
| --- | --- | ---: | ---: | ---: | ---: |
| Keynote-Large | ConvertSave | 745.53 | 555.32 | -25.5% | -3.2% |
| Keynote-Large | LoadProject | 15.63 | 18.60 | +19.0% | -20.3% |
| Keynote-Medium | ConvertSave | 25.32 | 26.78 | +5.8% | -8.3% |
| Keynote-Medium | LoadProject | 3.12 | 4.13 | +32.3% | -10.2% |
| Keynote-Small | ConvertSave | 6.64 | 8.14 | +22.5% | -2.6% |
| Keynote-Small | LoadProject | 2.43 | 2.67 | +9.8% | -2.3% |
| Numbers-Large | ConvertSave | 216.07 | 228.68 | +5.8% | -22.1% |
| Numbers-Large | LoadProject | 164.10 | 193.92 | +18.2% | -23.8% |
| Numbers-Medium | ConvertSave | 20.36 | 29.86 | +46.7% | -19.9% |
| Numbers-Medium | LoadProject | 14.33 | 14.40 | +0.5% | -21.8% |
| Numbers-Small | ConvertSave | 4.54 | 4.98 | +9.9% | -9.6% |
| Numbers-Small | LoadProject | 3.30 | 3.61 | +9.1% | -10.9% |
| Pages-Large | ConvertSave | 372.62 | 59.87 | -83.9% | +1.3% |
| Pages-Large | LoadProject | 12.61 | 13.74 | +8.9% | +0.2% |
| Pages-Medium | ConvertSave | 10.32 | 10.65 | +3.1% | +0.4% |
| Pages-Medium | LoadProject | 2.64 | 3.04 | +15.3% | -0.4% |
| Pages-Small | ConvertSave | 3.58 | 3.87 | +8.2% | -0.5% |
| Pages-Small | LoadProject | 2.45 | 2.71 | +10.9% | -0.8% |

## Limits and process memory

The host ran macOS 27.0.1 on Arm64 with ten logical processors, AC power and low
power mode disabled. Task builds and native test apps were idle. Unrelated host
processes remained active; snapshots included compiler and WindowServer activity,
about 22 GB used memory, 5 GB compressed and 1.3 GB free. Processor placement and
frequency were inherited. The slower cases remain visible, but this matrix does
not isolate whether elapsed differences come from the implementation or the host.

`AllocatedBytes` includes managed host invocation and concurrent process threads.
The collected heap observations in the raw samples include host/cache changes
after output validation and result release. A negative retained delta can reflect
collection of earlier state; it is not proof of leak freedom.

macOS `time -l` recorded whole-process maximum resident sizes of 404/368 MB for
the baseline matrices and 387/384 MB for the candidate matrices. Recorded peak
footprints were 340/339 MB and 324/325 MB respectively. These decimal values
include input construction, warmups, measurements, validation and the runtime.
They do not establish per-operation native or library-only peaks.

## Reproduction

Build each variant's workload from its selected checkout with the pinned SDK.
Keep its workload and complete owner assembly closure together. Use a shared
PowerForge source build containing operation allocation and collected managed
memory measurement. In a fresh PowerShell process for each matrix, run:

```powershell
./Build/Benchmarks/Run-IWorkRuntimeBenchmarks.ps1 `
    -BinaryRoot $binaryRoot -ModulePath $modulePath `
    -OutputRoot $outputRoot -WarmupCount 2 -IterationCount 5 -MeasureRetainedMemory
```

Choose explicit task output roots and alternate the two variants as above. Record
power, competing activity and processor placement before each run. On macOS,
wrapping each complete PowerShell invocation with `/usr/bin/time -l` records the
whole-process counters separately. Preserve failed samples before comparing
results. These synthetic inputs do not qualify large native Apple packages,
repeat-open stability, cancellation latency or portable ceilings; those remain
in [I5](../ROADMAP.md#i5-runtime-and-apple-host-acceptance).
