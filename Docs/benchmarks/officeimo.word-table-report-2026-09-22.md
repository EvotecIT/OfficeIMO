# Word structured-report cell writing (2026-09-22)

## Workload and change

The `WordReportOpenXmlEvidenceBenchmarks` workload creates a DOCX with a title,
summary, header, footer, and a two-column table containing 100 data rows. It
uses the public `WordDocument` and `WordTable.SetCellText` APIs and serializes
the completed document. The Open XML SDK validator checks the report values,
formatting, and Office-compatible package defaults before measurement.

`SetCellText` now fills the empty run already present in a newly created table
cell and walks cell siblings directly. Calls that replace existing content still
clear the first paragraph's content and preserve later paragraphs.

## Repeated isolated measurements

The table compares clean commit `b84f4670db61` with clean commit
`852f104512b7`. Each value is the median of three fresh-process probes from
`evidence --repeat 3` on .NET SDK 10.0.112. Allocation is per operation;
managed peak is the sampled growth over the probe's eight-operation batch.

| Host | Elapsed, baseline → change | Allocation, baseline → change | Managed peak, baseline → change | DOCX size, baseline → change |
| --- | ---: | ---: | ---: | ---: |
| Windows 11 build 26200, Ryzen 9 9950X3D2 | 3.93 → 3.93 ms | 1,075,228 → 1,008,969 B | 8,645,528 → 8,113,920 B | 19,856 → 19,855 B |
| Ubuntu 24.04.3 on WSL2, x64 | 4.87 → 4.32 ms | 1,069,499 → 990,979 B | 8,599,792 → 7,970,112 B | 20,255 → 20,257 B |
| macOS 27.0, Apple M4 | 6.13 → 3.55 ms | 1,056,292 → 990,078 B | 8,268,680 → 7,962,720 B | 19,852 → 19,852 B |

Allocation fell by 6.2% on Windows, 7.3% on Linux, and 6.3% on macOS. The
output sizes remained within three bytes of their respective baselines. These
short, cold child-process timings are sensitive to host load; the elapsed
observations are not a throughput gate or a claim of a stable speedup. The
Windows controlled Short BenchmarkDotNet run also had a wide confidence
interval, so it does not settle an elapsed-time budget.

Both commits passed `validate-openxml` on all three hosts. The focused
`SetCellText` tests passed on Windows for .NET 10, .NET 8, and .NET Framework
4.7.2, including replacement of populated cell content.

## Reproduce

From a clean checkout of either commit:

```text
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Word.Benchmarks -- validate-openxml
dotnet run --no-build -c Release -f net10.0 --project ./OfficeIMO.Word.Benchmarks -- evidence --repeat 3 --json .benchmark-artifacts/word/evidence.json
```

The runner also measures 100- and 1,000-item paragraph, report, read, and rich
replace lanes. This change reduces allocation in the report lane; stable
elapsed-time budgets and further reductions across the other Word workloads
remain on the [roadmap](../ROADMAP.md#document-format-depth).
