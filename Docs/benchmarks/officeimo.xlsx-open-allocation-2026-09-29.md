# XLSX reader opening allocation — 2026-09-29

The public `ExcelDocument.OpenDataReader` opening profile located the remaining allocation in shared-string and style loading. Its baseline on source commit `59847a7d9f83dd938a085da5e6f778901ea8e875` missed the 91,916-byte warmed allocation target for the 65K typed workbook scan. The bounded small shared-string change on commit `13838b2901de12e99440d22b288fa9d32d028272` reaches that fixture-specific target. These runs use .NET SDK 10.0.112 on Windows 11, Normal process priority, and the AMD Ryzen 9 9950X3D2's two 16-logical-processor affinity domains (`0xFFFF` and `0xFFFF0000`). The pinned 65,535-row, 14-column `65K_Records_Data.xlsx` fixture has SHA-256 `0F44D3E06454508DBD2CDBAF701B04160637162AB71471616D8ADC59D2EDD3A8`.

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

Raising the small-part pooled-read threshold from 4 KB to 32 KB alone did not materially reduce allocation on the same public scan. The [single-domain trial samples](xlsx-open-allocation-2026-09-29/trial-32k-a.csv) retain that rejected experiment. The final change instead pools parts only up to 8 KB and directly reads the narrow plain ASCII shared-string shape from that bounded buffer. XML markup, escaping, Unicode, alternate schemas, and whitespace-only text without `xml:space="preserve"` use the existing XML reader or package fallback. Focused tests cover the shared-string limits, escaped-value fallback, and the whitespace result across small and larger parts; the full .NET 8 Excel suite had 4,760 passes, five skips, and one desktop Excel COM smoke failure that passed in an isolated rerun.

For the final source, an adjacent control temporarily disabled only the small shared-string parser while retaining the 8 KB pooled-part threshold. The [control and final raw samples](xlsx-open-allocation-2026-09-29/) each contain 24 complete public scans per domain, including the first warm-up and all elapsed outliers. The table excludes only the first scan. Every scan returned checksum `4814962925905058108`.

| Variant | Domain | Warm median allocation | Warm median elapsed |
| --- | --- | ---: | ---: |
| XML control, 8 KB pooling | A | 106,704 B/scan | 75.43 ms/scan |
| Small ASCII table, final | A | 91,104 B/scan | 68.86 ms/scan |
| XML control, 8 KB pooling | B | 106,944 B/scan | 69.14 ms/scan |
| Small ASCII table, final | B | 91,424 B/scan | 78.04 ms/scan |

The final source is 50.3–50.4% below the original 183,832 B/scan allocation baseline on this workbook, and roughly 15.5 KB below its adjacent XML controls. The opposite elapsed-time directions across domains do not establish a speedup or slowdown. The fast path applies only to small, plain ASCII shared-string tables; other table shapes and large workbooks still need separate allocation and peak-memory qualification. Native Linux/macOS and portable elapsed-time budgets remain open.

To repeat the public profiles with the verified fixture, run `dotnet run -c Release -f net10.0 --project Docs/benchmarks/xlsx-metadata-2026-09-28/reproduction/Probe.csproj -- open-only <fixture-path> FFFF` and replace `open-only` with `typed-scan` or `sheet-names-100`. Use `FFFF0000` for the other processor domain after checking the current machine topology. Run each profile in a fresh process and retain the first warm-up separately.
