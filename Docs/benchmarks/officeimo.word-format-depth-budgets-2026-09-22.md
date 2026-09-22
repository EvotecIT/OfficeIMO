# Word 100-item format-depth budgets (2026-09-22)

## Workloads and proof

The opt-in Word evidence runner now checks absolute budgets for three complete
OfficeIMO workflows: create and serialize 100 paragraphs, create and serialize
the 100-row structured report, and load, replace, and serialize 100 rich
paragraphs. A sample output from each isolated probe is reopened and checked
for its text, structure,
formatting, and package defaults. The report includes title, summary, header,
footer, and every table cell; replacement preserves the input style catalog.

The validator also runs equivalent Open XML SDK operations over the same rich
package shell. That comparison establishes output validity; the checked-in
budgets apply to the OfficeIMO operations only. This benchmark project remains
outside the normal solution and carries no new Word runtime dependency.

## Three-host baseline

Each host measured clean commit `727eaf53e9e5256960ef3a3121730290354ac0ba`
with .NET SDK 10.0.112 and three isolated child probes per workload. The table
reports medians for the 100-item OfficeIMO operations.

| Host | Workload | Elapsed ms/op | Allocated bytes/op | Managed-peak growth bytes | DOCX bytes |
| --- | --- | ---: | ---: | ---: | ---: |
| Windows 11, Ryzen 9 9950X3D2 | Paragraph create | 2.68 | 636,794 | 5,129,248 | 18,704 |
| Windows 11, Ryzen 9 9950X3D2 | Structured report | 4.36 | 991,537 | 7,974,440 | 19,856 |
| Windows 11, Ryzen 9 9950X3D2 | Rich replace | 2.00 | 750,627 | 6,042,744 | 18,655 |
| Ubuntu 24.04.3, x64 WSL2 | Paragraph create | 1.90 | 619,261 | 4,989,056 | 19,090 |
| Ubuntu 24.04.3, x64 WSL2 | Structured report | 4.28 | 990,978 | 7,970,080 | 20,255 |
| Ubuntu 24.04.3, x64 WSL2 | Rich replace | 5.23 | 751,580 | 6,050,488 | 19,040 |
| macOS 27.0, Apple M4 | Paragraph create | 1.50 | 618,083 | 4,979,448 | 18,699 |
| macOS 27.0, Apple M4 | Structured report | 5.09 | 990,036 | 7,962,720 | 19,851 |
| macOS 27.0, Apple M4 | Rich replace | 6.65 | 749,164 | 6,023,400 | 18,653 |

The macOS runner returned zero for process working-set peak, so that metric is
unavailable there. The observed Windows process-peak medians ranged from 82.6
to 91.4 MiB; Ubuntu medians ranged from 103.6 to 119.1 MiB. Child-process
elapsed time varied substantially, especially on Ubuntu and macOS, so these
medians do not establish a speedup.

## Checked-in limits

`OfficeIMO.Word.Benchmarks/word-format-depth-budgets.json` sets absolute
regression limits for every isolated OfficeIMO probe, with headroom above the
largest observed sample. These limits are a guard for future work; they are not
a claim that the remaining allocation and timing costs are optimal.

| Workload | Max elapsed ms/op | Max allocation bytes/op | Max managed-peak growth | Max process peak | Max DOCX bytes |
| --- | ---: | ---: | ---: | ---: | ---: |
| Paragraph create | 10 | 700,000 | 6 MiB | 140 MiB | 20,000 |
| Structured report | 20 | 1,100,000 | 9 MiB | 140 MiB | 21,000 |
| Rich replace | 20 | 830,000 | 7 MiB | 140 MiB | 20,000 |

The process limit is observable on Windows and Ubuntu. On macOS, the runner
reports that this one limit could not be observed while still checking the
other four metrics. Elapsed limits allow the variation seen in isolated runs;
controlled timing remains necessary before claiming throughput improvements.

## Reproduce

From a clean checkout with the pinned SDK:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Word.Benchmarks -- validate-openxml
dotnet run -c Release -f net10.0 --no-build --project .\OfficeIMO.Word.Benchmarks -- evidence --repeat 3 --json .\.benchmark-artifacts\word\evidence.json --budget .\OfficeIMO.Word.Benchmarks\word-format-depth-budgets.json
```

The earlier [structured-report cell-writing result](officeimo.word-table-report-2026-09-22.md)
records the delivered allocation reduction. Further reductions and stable
elapsed-time proof across these three workflows remain open on the
[roadmap](../ROADMAP.md#document-format-depth).
