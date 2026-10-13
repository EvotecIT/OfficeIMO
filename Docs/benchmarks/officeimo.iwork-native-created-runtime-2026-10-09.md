# Native Keynote scale observations — 2026-10-09

The OfficeIMO writer produced a 1,000-slide Keynote package with unique text on
every slide. All 18 measured samples passed complete semantic validation across
source reading, editable PPTX conversion/save, and native Keynote encoding/save.
This extends the [runtime lane](../../OfficeIMO.IWork.Benchmarks/README.md) with
an owned native package. Large independent Apple documents and native application
acceptance remain separate qualification criteria.

The [raw samples, summaries and assembly hashes](iwork-native-created-runtime-2026-10-09.json)
identify the measured candidate. OfficeIMO assemblies were built from base
`d0301ab016e0c1d3d15834f311a97268a101f0ac` with the fidelity/workflow source changes
identified by those hashes. The PowerForge source owner is
`54a8a91706b9406529884f28ab63897399c28be4`.

## Work and host

Each native input is 1,209,406 bytes. Input/model construction and all semantic
checks stay outside measurement. Reading validates every slide and text value.
PPTX conversion/save requires complete editable reconstruction and reopens the
saved presentation. Native saving encodes the prepared model and reopens the
saved Keynote package. Both checks retain the expected paragraph terminators.
The native inputs have matching SHA-256 hashes across runs, and every output is
checked before a sample succeeds. Semantic readback and native writer determinism
have separate contracts.

Two fresh PowerShell 7.5.4 processes ran on .NET 9.0.10, loading the `net8.0`
workload and owner assemblies built with OfficeIMO SDK 10.0.112. The host was
EvoMini, macOS 27.0.1, arm64, with ten logical processors. Each run used one
warmup and three measured iterations per operation, rotated order, collection
before iterations and no outlier removal. Other source builds and tests were
active. Processor placement was inherited. These observations include contention
and do not establish performance changes or portable budgets.

| Operation | Standard median ms | Mean allocated MiB | Standard elapsed range ms |
| --- | ---: | ---: | ---: |
| Read/project | 209.77 | 112.31 | 177.08–850.14 |
| Convert/save PPTX | 3,460.94 | 628.96 | 2,211.48–7,046.36 |
| Encode/save Keynote | 72.99 | 94.29 | 60.55–279.13 |

Allocation includes managed host invocation and concurrent process threads.
Collected managed-heap observations include host and cache state after semantic
validation and result release. Negative retained deltas reflect collected state;
they do not establish leak freedom or native allocation behavior.

The second run sampled managed heap and resident memory at a requested 5 ms
interval. Every sample retained valid observations, with no sampling failures.
Its timing is separate from the standard run. Sampled maxima are lower bounds on
peaks, include observer/host effects and do not attribute native memory to a
library. The raw samples retain sample counts, baselines, maxima and deltas.

## Reproduction

Build the workload and PowerForge module for `net8.0` with each repository's
pinned SDK. Keep the matching owner assemblies together. In a fresh PowerShell
process, run:

```powershell
./Build/Benchmarks/Run-IWorkRuntimeBenchmarks.ps1 `
    -BinaryRoot $binaryRoot -ModulePath $modulePath -OutputRoot $outputRoot `
    -Scale NativeCreated -Kind Keynote -Operation LoadProject,ConvertSave,NativeSave `
    -WarmupCount 1 -IterationCount 3 -MeasureRetainedMemory
```

Use another fresh process with `-MemorySamplingIntervalMilliseconds 5` for the
instrumented run. Keep competing work idle when collecting acceptance budgets.
These operations do not qualify mid-projection, destination-construction or
native-encoding cancellation; the existing active I/O and workflow-publication
checks have their own evidence.
