# XLSX reader opening allocation — 2026-09-29

The public `ExcelDocument.OpenDataReader` path still misses the 91,916-byte warmed allocation target for the 65K typed workbook scan. This run separates reader opening from row traversal on source commit `59847a7d9f83dd938a085da5e6f778901ea8e875`. It uses .NET SDK 10.0.112 on Windows 11, Normal process priority, and the AMD Ryzen 9 9950X3D2's two 16-logical-processor affinity domains (`0xFFFF` and `0xFFFF0000`). The pinned 65,535-row, 14-column `65K_Records_Data.xlsx` fixture has SHA-256 `0F44D3E06454508DBD2CDBAF701B04160637162AB71471616D8ADC59D2EDD3A8`.

The [probe](xlsx-metadata-2026-09-28/reproduction/Program.cs) measures complete public calls, including reader disposal. `open-only` validates the 14-column schema without advancing a row. `typed-scan` checks the row count and typed-value checksum. Those profiles make 24 calls per process; `sheet-names-100` makes 21 batches of 100 calls. The first call or batch is excluded as warm-up. The [raw samples](xlsx-open-allocation-2026-09-29/) retain all observations and outliers. The sheet-name result is per 100 calls and is divided by 100 for the allocation column below.

| Profile | Domain | Warm median allocation | Warm median elapsed |
| --- | --- | ---: | ---: |
| Open only | A | 106,592 B/open | 50.92 ms/open |
| Open only | B | 106,592 B/open | 43.88 ms/open |
| Typed 65K scan | A | 106,824 B/scan | 83.89 ms/scan |
| Typed 65K scan | B | 106,920 B/scan | 72.57 ms/scan |
| Sheet names | A | 51,649 B/open | 16.96 ms/100 opens |

Both full-scan domains produced checksum `4814962925905058108`. For this fixture, nearly all measured managed allocation occurs while opening the reader; traversal adds little. This does not establish that arbitrary workbooks or caller result materialization are allocation-free. The elapsed results are single process runs and are not a speed comparison or a portable budget.

Temporary `GC.GetAllocatedBytesForCurrentThread` checkpoints, removed after profiling, placed about 45.2 KB of each warm open inside indexed-row validation: loading 224 shared strings accounted for about 27.4 KB and loading styles about 17.8 KB. The shared-string XML reader setup accounted for about 13.8 KB of its total. Earlier setup stages included about 9.4 KB for content types, 7.6 KB for the root relationship and workbook validation, 7.4 KB for workbook relationships, and 11.9 KB for workbook XML. These internal figures are diagnostic attribution, not a separate public benchmark; instrumentation and its output were absent from the public samples above.

Raising the small-part pooled-read threshold from 4 KB to 32 KB did not materially reduce allocation on the same public scan. The [single-domain trial samples](xlsx-open-allocation-2026-09-29/trial-32k-a.csv) retain that rejected experiment; the production threshold remains 4 KB. Further work should measure a bounded parser change against the public calls and preserve XML validation, strict/transitional namespaces, malformed-package fallback, metadata limits, cancellation, and pool ownership. The 50% reduction target relative to the original 183,832 B/scan remains open, as do native Linux/macOS and peak-memory budgets.

To repeat the public profiles with the verified fixture, run `dotnet run -c Release -f net10.0 --project Docs/benchmarks/xlsx-metadata-2026-09-28/reproduction/Probe.csproj -- open-only <fixture-path> FFFF` and replace `open-only` with `typed-scan` or `sheet-names-100`. Use `FFFF0000` for the other processor domain after checking the current machine topology. Run each profile in a fresh process and retain the first warm-up separately.
