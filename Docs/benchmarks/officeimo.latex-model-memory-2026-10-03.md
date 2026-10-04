# LaTeX model memory evidence (2026-10-03)

Compact internal token records reduce the memory used by lossless parsing without
requiring callers to inspect every public token. The large Windows corpus allocates
43.42–43.82 MiB and retains 34.63 MiB, compared with 53.30 MiB allocated and 42.54 MiB
retained before this change. Complete public-token inspection costs more allocation
and retains almost the same heap as before. These results qualify specific workloads;
they do not establish a universal throughput improvement.

The [measurement record](officeimo.latex-model-memory-2026-10-03.json) contains both
run orders, all 288 retained BenchmarkDotNet timing samples, fresh-process memory
ranges, platform observations, input hashes, source refs and loaded assembly hashes.
The [benchmark README](../../OfficeIMO.Latex.Benchmarks/README.md) describes how to
validate and run these opt-in lanes.

## Workloads and controls

All lanes build the complete lossless syntax and semantic model. Setup checks exact
source recovery, diagnostics, headings and content markers. Preserve writing must
reproduce the source exactly. `ParseInspect` also reads and validates every public
token's text, value, span and termination state, retaining the model afterward.

| Corpus | UTF-8 bytes | Tokens |
| --- | ---: | ---: |
| Small | 1,470 | 548 |
| Normal | 68,692 | 25,636 |
| Large | 701,399 | 256,036 |

Inputs use fixed CRLF line endings and matching hashes on Windows, Linux and macOS.
The large input SHA-256 is
`36470963F33460A48CDC314D02ACBDDF23784402B78F4D535B757F603946AF0B`.

The Windows comparison uses .NET 10.0.12 and BenchmarkDotNet 0.15.8 on an AMD Ryzen
9 9950X3D2 with 16 cores and 32 logical processors. Binaries are built with pinned
SDK 10.0.112; the isolated in-process benchmark host reports the machine's SDK
10.0.400, without rebuilding those binaries. Processes use High
priority and affinity mask 1. Four runs execute strictly sequentially: before,
after, after, before. Each run has three fresh child processes per memory case,
then three warmups and eight measured iterations per BenchmarkDotNet case using
the in-process emit toolchain. Outliers are retained. Other host work was active;
the preflight total CPU sample was 41%, and elapsed results show substantial noise.

The baseline runtime is `0ae340322bfcd532cf2caebad6f86dba996b72a4`; candidate runtime
is `dfe7fae293d0ef966cf3cc876e33a93b16e371f2`. Both use the measurement harness from
`c7afeb960338781e31acc99a90a4f47ecf5071eb` and identical Core and harness binaries.
Each measurement validates the loaded native/harness hashes. The isolated Windows
binary directories contain no Git checkout, so their automatic source-ref fields
are `unknown`; the explicit source/hash manifest supplies that provenance.

## Allocation and retained heap

The table gives the range of the per-run medians from the two controlled orders.
Retained heap is measured while keeping the parsed document alive. Parse plus write
keeps only the returned string, so its retained heap is a different contract.

| Workflow | Before allocated MiB | After allocated MiB | Before retained MiB | After retained MiB |
| --- | ---: | ---: | ---: | ---: |
| Small parse | 0.132–0.138 | 0.137–0.138 | 0.077–0.080 | 0.067 |
| Normal parse | 5.199 | 4.151–4.185 | 4.215 | 3.392–3.393 |
| Large parse | 53.300 | 43.418–43.818 | 42.539 | 34.627 |
| Large parse + write | 53.339 | 43.456–43.857 | about 0 | about 0 |
| Small parse + full token inspection | 0.148–0.154 | 0.182–0.183 | 0.093–0.096 | 0.089 |
| Normal parse + full token inspection | 5.935 | 6.275–6.309 | 4.951 | 4.907 |
| Large parse + full token inspection | 60.658 | 64.638–65.037 | 49.895 | 49.957 |

Large parsing reduces allocation by approximately 18% and retained heap by 18.6%.
The full-inspection lane increases large allocation by approximately 6.6–7.2%,
with retained heap about 0.1% higher. Steady-state small parsing allocates 142,065 B
versus 135,952–135,953 B, approximately 4.5% more. Consumers inspecting every token
pay for materializing the public objects; compact records are released after each
fully materialized chunk. Sparse inspection preserves stable public object identity.

On Windows, the large parse's sampled managed-heap estimate peak falls from 53.43
MiB to 38.83–40.77 MiB; reported process peaks fall from 93.54–93.93 MiB to
81.52–83.37 MiB. Full inspection has mixed sampled peaks and a smaller process-peak
reduction. These measurements include the child-process and sampling boundaries
described below.

## Elapsed results

| Workflow | Forward before mean ms | Forward after mean ms | Reverse before mean ms | Reverse after mean ms |
| --- | ---: | ---: | ---: | ---: |
| Small parse | 0.071 | 0.071 | 0.062 | 0.066 |
| Normal parse | 8.569 | 6.101 | 6.782 | 6.267 |
| Large parse | 189.395 | 265.050 | 211.440 | 180.002 |
| Large parse + write | 137.427 | 139.711 | 97.852 | 91.833 |
| Large parse + full token inspection | 189.334 | 132.139 | 126.441 | 122.421 |

Large parse medians are 217.67/261.81 ms before/after in the forward order and
217.40/214.56 ms in the reverse order. The opposing order effects and wide sample
ranges prevent a general speed claim. Memory reduction is the demonstrated result.

## Platform and measurement boundaries

Actual native builds and focused LaTeX, conversion and Reader suites pass on Windows
(.NET 8, .NET 10 and .NET Framework 4.7.2), Linux/WSL x64 (.NET 8 and .NET 10), and
macOS Arm64 (.NET 8 and .NET 10). Public Reader path, stream and byte-array routes
preserve the same text hashes for the pinned official LaTeX2e inputs, block/document
chunking and requested size limits. Linux source is on an NTFS-mounted WSL path;
this does not qualify a case-sensitive ext4 checkout or every Reader format.

| Candidate large parse, fresh child | Allocated MiB | Retained MiB | Process peak |
| --- | ---: | ---: | --- |
| Windows x64 | 43.42–43.82 | 34.63 | 81.52–83.37 MiB |
| Linux/WSL x64 | 43.42 | 34.63 | Recorded in the measurement record |
| macOS Arm64 | 43.42 | 34.63 | Unavailable through the measured BCL API |

Retained managed growth uses `HeapSizeBytes - FragmentedBytes` from the last forced
full blocking collection, minus the same baseline metric. Microsoft documents
[GC heap information](https://learn.microsoft.com/en-us/dotnet/api/system.gcmemoryinfo?view=net-10.0)
and the [heap-size](https://learn.microsoft.com/en-us/dotnet/api/system.gcmemoryinfo.heapsizebytes?view=net-10.0)
and [fragmentation](https://learn.microsoft.com/en-us/dotnet/api/system.gcmemoryinfo.fragmentedbytes?view=net-10.0)
fields used for this boundary. The older `GC.GetTotalMemory` retained estimate is
kept separately; macOS observations show it can differ materially from the post-GC
live heap, so it is not used for the retained-graph claim.

The managed peak is a 1 ms sampled `GC.GetTotalMemory` estimate, not a precise heap
high-water mark. Its macOS values are substantially higher than Windows/Linux and
must not be substituted for retained heap or precise native/process memory.
Missing process peaks are recorded as `null`; budget verification explicitly fails
instead of interpreting them as zero. An independent macOS `/usr/bin/time -l`
observation reports 116,883,456 B maximum resident size for one complete child
lifecycle; this is a separate capture boundary from the BCL sample. Full portable
process-peak and timing-budget qualification remains open.

Timing and memory ceilings remain opt-in, outside correctness CI. The existing
allocation and retained ceilings are retained; these runs do not establish that
every host-dependent timing or peak budget passes on every platform. Neither
external TeX compilation nor benchmark tooling enters the shipped runtime graph.
