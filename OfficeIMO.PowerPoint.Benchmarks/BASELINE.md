# PowerPoint package-workflow evidence — 2026-09-22

This evidence covers package create/save and open/edit/save only. Image and PDF
export have separate owners and are intentionally excluded. The original
comparison implementation is `6ff903bb2`; the current large open/edit/save
reduction is `950f0fc2c7`. The benchmark runner and comparison projects are
opt-in and keep third-party dependencies outside the normal solution.

The comparison policy is strict: a lane is a contender only when both elapsed
time and managed allocation are no more than 2× the equivalent implementation.
A result above 2× through 5× is a material remediation gap, more than 5× is
unacceptable unless the contracts differ, and 40× is only an incident threshold.

## Contract and method

Both implementations create or open editable presentations with the same slide
dimensions, background and style pattern, text, vector panels, tables,
two-series charts, and every-tenth-slide edit cadence. The open/edit/save lane
uses the exact same OfficeIMO-authored input bytes. Shape counts are not compared
because the APIs expose compound table and chart content differently.

Every sample runs in a fresh process. Timing, allocation, sampled managed-heap
growth, process peak working set, input bytes, and output bytes are captured
before validation. The resulting package is then reopened and must pass the
shared semantic checks for expected text, styling, table contents, chart data,
edit markers, slide count, and Open XML validity.

Windows results use .NET 8.0.30 on Windows 10.0.26200 x64 with five samples per
lane. Linux results use .NET 8.0.30 on Ubuntu 24.04 x64 under WSL with three
samples per lane. Tables report medians; ratios are OfficeIMO divided by
ShapeCrawler 0.79.4.

## Windows medians

| Workflow | Scale | OfficeIMO ms | ShapeCrawler ms | Time ratio | OfficeIMO alloc MiB | ShapeCrawler alloc MiB | Allocation ratio | Peak ratio | Output-size ratio |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Create/save | Small | 313.68 | 270.85 | 1.16× | 8.17 | 10.35 | 0.79× | 0.98× | 0.850× |
| Create/save | Normal | 355.72 | 337.97 | 1.05× | 16.14 | 48.73 | 0.33× | 0.71× | 0.965× |
| Create/save | Large | 425.87 | 681.97 | 0.62× | 47.02 | 189.35 | 0.25× | 0.87× | 1.027× |
| Open/edit/save | Small | 194.23 | 210.18 | 0.92× | 5.28 | 6.44 | 0.82× | 1.00× | 0.996× |
| Open/edit/save | Normal | 269.26 | 210.99 | 1.28× | 13.30 | 11.46 | 1.16× | 1.04× | 0.996× |
| Open/edit/save | Large | 387.53 | 284.62 | 1.36× | 42.90 | 31.28 | 1.37× | 1.15× | 0.995× |

All six package lanes are inside the 2× contender ceiling for elapsed time and
allocation. The worst current margins are 1.36× elapsed and 1.37× allocation;
three lanes allocate less than ShapeCrawler, and the large create/save lane is
faster while allocating one quarter as much managed memory.

## Linux medians

| Workflow | Scale | OfficeIMO ms | ShapeCrawler ms | Time ratio | Allocation ratio | Peak ratio |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| Create/save | Small | 454.62 | 464.31 | 0.98× | 0.79× | 0.98× |
| Create/save | Normal | 512.55 | 865.59 | 0.59× | 0.33× | 0.74× |
| Create/save | Large | 559.82 | 1559.30 | 0.36× | 0.25× | 0.87× |
| Open/edit/save | Small | 254.81 | 339.84 | 0.75× | 0.82× | 1.02× |
| Open/edit/save | Normal | 342.15 | 359.01 | 0.95× | 1.17× | 1.05× |
| Open/edit/save | Large | 498.25 | 465.47 | 1.07× | 1.38× | 1.14× |

The refreshed cross-platform sample preserves the same classification. The
worst Linux margin is 1.38× allocation; every elapsed-time lane is at or below
1.07×.

## Optimization and compatibility boundary

The first dominant large-deck cost was save-time cloning of a package that had
already been saved into the presentation's owned memory stream. For ordinary
macro-free PPTX output on .NET 8 and later, OfficeIMO snapshots that finalized
stream directly. Package conversion and VBA-preserving paths continue to clone.

A second full-package clone remained in the signature mutation policy: every
ordinary unsigned save serialized the whole package solely to prove that no
signature carrier existed. The current path inspects the bounded package bytes
first and returns immediately when unsigned. Live or malformed signature
carriers, and package implementations whose current stream is not parseable,
still take the original fail-closed full snapshot path.

The large Windows open/edit/save allocation fell from about 74.1 MiB to
56.9 MiB after the first change and to 42.9 MiB now. On the same machine, the
remaining signature change reduced its save stage from 27.4 MiB to 13.4 MiB;
large create/save fell from 61.8 MiB to 47.0 MiB. .NET
Framework 4.7.2 requires the clone path because its packaging implementation
does not finalize compressed-part lengths early enough for a safe live-stream
snapshot. The full PowerPoint suite passes on net472, net8.0, and net10.0 with
that boundary.

## Large open/edit/save reduction — 2026-09-22

Opening a presentation previously materialized every slide's shape tree before
the caller accessed or edited it. Saving also loaded every untouched slide to
count hidden slides and rewrite its root. The current path defers shape wrappers
until a shape operation needs them, preserves untouched slide XML, and reads the
root `show` attribute through the bounded XML reader when hidden-slide metadata
is refreshed. Loaded and legacy-visibility slides still use the normal save and
normalization paths; notes retain their separate save path.

The measurements compare clean exact commits `0f8249112d` and `950f0fc2c7`
with .NET 10.0.12. Each cell is the median of five isolated child processes for
the 120-slide, 1,492-shape corpus. Every output reopened, passed the semantic
oracle, and passed Open XML validation.

| Host | Base ms | Current ms | Current allocation MiB | Current managed peak MiB | Current process peak MiB | Current output bytes |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Windows 10 x64 | 411.42 | 272.41 | 25.86 | 25.95 | 81.07 | 339,911 |
| Ubuntu 24.04 x64 | 496.25 | 485.48 | 25.85 | 25.94 | 100.86 | 338,010 |
| macOS 27 Apple M4 | 517.62 | 226.19 | 25.88 | 19.04 | unavailable | 340,149 |

The base allocation was 42.0 MiB on all three hosts. Current allocation is
38.5% lower. Managed peak fell 38.5% on Windows and Linux and 53.1% on macOS;
process peak fell 16.7% on Windows and 13.9% on Linux. macOS returned zero for
the process peak counter, so that metric is unavailable there. Output size is
unchanged within normal ZIP variation.

Elapsed time improved 33.8% on Windows, 2.2% on the selected Linux repeat, and
56.3% on Apple M4. The Linux candidate samples ranged from 294 to 637 ms while
allocations stayed within 0.1 MiB, so the Linux timing is directional rather
than a stable throughput claim.

## Regression gates

`powerpoint-performance-budgets.json` covers all six package lanes. The large
open/edit/save ceilings are now 1,500 ms, 36 MiB allocation, 40 MiB managed-heap
growth, 192 MiB process peak, and 352 KiB output. Allocation, managed peak, and
output size are hard ceilings. Elapsed time and process peak retain wider
headroom to catch gross stalls and runaway memory without turning workstation
noise into a throughput claim.

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.PowerPoint.Benchmarks -- --verify-budgets
```

Use repeated same-machine comparisons for smaller timing changes. The checked-in
budget is a regression guard, not a throughput guarantee.
