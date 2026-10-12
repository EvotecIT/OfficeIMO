# iWork phase cancellation observations — 2026-10-11

All twelve measured samples observed cancellation while projecting a loaded
1,000-slide source or encoding its native Keynote model before caller-stream
copying. Every sample retained the source content hash, left the sentinel
destination untouched and writable, and recovered all 1,000 ordered slide text
values through a subsequent uncancelled projection.

The [samples, summaries and assembly hashes](iwork-phase-cancellation-2026-10-11.json)
identify source candidate `2604a7c50052b01d20deb63cc8f95207b19f24f5` and PowerForge
source `54a8a91706b9406529884f28ab63897399c28be4`. The input is an OfficeIMO-created
1,209,418-byte package, SHA-256
`CE2D0BF12301FEDDED7294815E77392D0C90F357A1CE5041BF9FFB69639936BD`.
This is owned-producer scale evidence.

## Observation boundary

A timer requests cancellation after 25 ms. A successful sample requires at
least 20 ms elapsed and 5 ms of process CPU advancement before the request.
Observed pre-request elapsed time was 26.63–94.53 ms. Process CPU includes
observer and host activity; timer scheduling does not identify an internal
parser or encoder checkpoint. An untouched destination distinguishes native
encoding cancellation from stream-copy cancellation. The timer callback is
joined before its token source and process handle are released.

Request-to-exception latency ranged from 0.098 to 12.140 ms for projection and
0.051 to 65.561 ms for native encoding. Unrelated Swift and .NET builds were
active on the host. These contended observations do not set latency budgets or
demonstrate a performance change.

Two fresh PowerShell 7.5.4 processes ran on .NET 9.0.10 with the `net8.0`
workload assemblies built using OfficeIMO SDK 10.0.112. The Mac16,10 host has an
Apple M4, ten physical/logical processors, and macOS 27.0.1 build 26A434. Low
power mode was disabled; processor placement was inherited and unqualified.
`DOTNET_PROCESSOR_COUNT=2` applied to each host. Each pass used one warmup and
three measured iterations per operation, rotated ordering, collection before
iterations, and no outlier removal. PowerForge owns the measurements; preparation,
source hashing and complete recovery checks stay outside timing.

The standalone correctness command passed all sixteen existing operations on
.NET 8 and .NET 10. The separate opt-in phase command passed both operations on
each target. A matching prefix of 999 slides was rejected by the recovery
cardinality guard.

The separate native Numbers smoke validated all nine saved typed cells. Its
explicit partial policy accepts the pinned fixture's reported padding omission
and retains automatic-decimal and border-width approximations. Preview fallback,
unrelated losses and unresolved selected content fail validation. This benchmark
policy does not change production conversion or complete PDF acceptance.

## Reproduction and limits

Build the workload and PowerForge module with their pinned SDKs as described in
the [runtime lane](../../OfficeIMO.IWork.Benchmarks/README.md). In two fresh
PowerShell processes, run:

```powershell
./Build/Benchmarks/Run-IWorkRuntimeBenchmarks.ps1 `
    -BinaryRoot $binaryRoot -ModulePath $modulePath -OutputRoot $outputRoot `
    -Scale NativeCreated -Kind Keynote `
    -Operation CancelDuringProjection,CancelDuringNativeEncoding `
    -WarmupCount 1 -IterationCount 3
```

Keep competing work idle before selecting host-specific budgets. These
timer-dependent cases remain outside ordinary correctness CI. They do not
qualify large independent Apple documents, destination construction, atomic
path staging, other platforms, peak/native-memory retention, or physical Apple
host acceptance.
