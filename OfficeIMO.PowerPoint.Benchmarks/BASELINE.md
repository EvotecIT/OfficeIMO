# PowerPoint package-workflow evidence — 2026-08-24

This evidence covers package create/save and open/edit/save only. Image and PDF
export have separate owners and are intentionally excluded. The production
candidate is `6ff903bb2`; the benchmark runner and comparison projects are
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

## Regression gates

`powerpoint-performance-budgets.json` covers all six package lanes. Allocation
ceilings are now 8-62 MiB and managed-heap-growth ceilings are 16-64 MiB, with
roughly 30% headroom on the normal and large current measurements. Output size
is also a hard ceiling. Elapsed time and process peak use wider ceilings to
catch gross stalls and runaway memory without turning workstation noise into a
throughput claim.

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.PowerPoint.Benchmarks -- --verify-budgets
```

Use repeated same-machine comparisons for smaller timing changes. The checked-in
budget is a regression guard, not a claim that the remaining 1.36-1.38× margins
cannot improve.
