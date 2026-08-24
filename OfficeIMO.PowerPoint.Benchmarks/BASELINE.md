# PowerPoint package-workflow evidence — 2026-08-24

This evidence covers package create/save and open/edit/save only. Image and PDF
export have separate owners and are intentionally excluded. The production
candidate is `47365679c`; the benchmark runner and comparison projects are
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
| Create/save | Small | 288.36 | 248.54 | 1.16× | 9.79 | 10.35 | 0.95× | 1.00× | 0.850× |
| Create/save | Normal | 379.81 | 337.02 | 1.13× | 20.61 | 48.73 | 0.42× | 0.75× | 0.965× |
| Create/save | Large | 546.45 | 684.81 | 0.80× | 61.80 | 189.44 | 0.33× | 0.90× | 1.027× |
| Open/edit/save | Small | 193.94 | 198.23 | 0.98× | 6.88 | 6.44 | 1.07× | 1.02× | 0.996× |
| Open/edit/save | Normal | 245.41 | 209.26 | 1.17× | 17.59 | 11.46 | 1.53× | 1.11× | 0.996× |
| Open/edit/save | Large | 440.17 | 259.06 | 1.70× | 56.93 | 31.28 | 1.82× | 1.34× | 0.995× |

All six package lanes are inside the 2× contender ceiling for both elapsed time
and allocation. The large open/edit/save lane is close enough to the ceiling to
remain an optimization target; it is not treated as a comfortable margin.

## Linux medians

| Workflow | Scale | OfficeIMO ms | ShapeCrawler ms | Time ratio | Allocation ratio | Peak ratio |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| Create/save | Small | 432.07 | 442.93 | 0.98× | 0.95× | 0.99× |
| Create/save | Normal | 445.31 | 656.68 | 0.68× | 0.42× | 0.78× |
| Create/save | Large | 661.80 | 1448.40 | 0.46× | 0.32× | 0.96× |
| Open/edit/save | Small | 275.99 | 384.44 | 0.72× | 1.07× | 1.04× |
| Open/edit/save | Normal | 335.91 | 455.23 | 0.74× | 1.53× | 1.10× |
| Open/edit/save | Large | 623.07 | 385.10 | 1.62× | 1.82× | 1.30× |

The cross-platform sample preserves the same classification. Allocation ratios
are especially stable across Windows and Linux.

## Optimization and compatibility boundary

The dominant large-deck cost was save-time cloning of a package that had already
been saved into the presentation's owned memory stream. For ordinary macro-free
PPTX output on .NET 8 and later, OfficeIMO now snapshots that finalized stream
directly. Package conversion and VBA-preserving paths continue to clone.

The large Windows open/edit/save allocation fell from about 74.1 MiB to
56.9 MiB, and its save stage fell from about 46.9 MiB to 28.7 MiB. .NET
Framework 4.7.2 requires the clone path because its packaging implementation
does not finalize compressed-part lengths early enough for a safe live-stream
snapshot. The full PowerPoint suite passes on net472, net8.0, and net10.0 with
that boundary.

## Regression gates

`powerpoint-performance-budgets.json` covers all six package lanes. Allocation,
sampled managed-heap growth, and output size are hard ceilings. Elapsed time and
process peak use wider ceilings to catch gross stalls and runaway memory without
turning workstation noise into a throughput claim.

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.PowerPoint.Benchmarks -- --verify-budgets
```

Use repeated same-machine comparisons for smaller timing changes. The checked-in
budget is a regression guard, not proof that the large edit lane needs no more
work.
