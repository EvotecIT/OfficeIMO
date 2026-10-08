# Independent tabular comparisons

This opt-in .NET 10 project compares workbook reads and writes using pinned
generated and real-data workloads. CSV and Arrow comparisons are explicit build
options. It stays outside the normal solution and adds no runtime dependencies
to OfficeIMO. BenchmarkDotNet 0.15.8 controls the measurements.

The [8 October 2026 report](../../Docs/benchmarks/officeimo.excel-tabular-2026-10-08.md)
records the captured Windows .NET 10 matrix, artifact and input identities,
measurement samples, warnings, and first-row versus full-scan boundaries.

The default input has one header and 50,000 records with `Name`, `Id`, `Date`, and
`Value`. Names repeat from eight ASCII values, IDs run from 1 through 50,000, dates
use `DateTime.FromOADate(45292 + index % 3650 + 0.25)`, and values use `index * 1.5`.
Generation uses the same `XlsxSheetWriter.WriteRecordsAsync` operation and
default options as the [upstream typed workload](https://github.com/GabrielMarquezMatte/ExcelReader/blob/ca5b50f99e8ef57ab476f0a2bc8043558d58b28d/tests/ExcelReader.Benchmarks/Shared/WorkbookGenerator.cs).
The original worksheet has inline strings, no dimension, and no row or cell
references. Setup rejects a generated input that differs from that shape.

Every typed XLSX scan opens and disposes the reader, materializes the same four-property
record, consumes all fields, and verifies the row count and the upstream
accumulator. Setup additionally compares every field in every row, checks the
DataReader headers and single-sheet count, and prints package/XML size and
coordinate counts. Fixture generation and validation are outside the timed work.

Setup prints the SHA-256 and length of the actual generated typed/raw workbook
bytes, string-heavy and cold inputs, Arrow conversion sources, and UTF-8/Arrow
writer buffers. Validated writer packages are labelled as outputs. Prepared
object lists have a generator and complete field proof rather than an input-file
hash. ZIP fixtures also log each entry's name, uncompressed length, payload hash,
compressed length, and timestamp. The producer and diagnostic rewrites use ZIP
creation timestamps, so separately generated package hashes can differ. Matching
entry payload hashes distinguish container metadata or compression changes from
changed document parts; do not infer equal payloads from equal field checksums.
All implementations within an individual read case consume the same byte array.
Cold identity logging hashes package bytes alone and does not open ZIP entries.
The existing cold fixture shape check stays in setup; full field proof runs
through the separate `--validate-cold` qualification command.

OfficeIMO uses `ExcelDocument.OpenDataReader` and typed getters; ExcelReader uses
`ExcelParser.FromAttributes<TypedRecord>().Parse(workbook.FirstSheet)`; Sylvan
uses `GetRecords<TypedRecord>()`. These reproduce the public API work in the
[upstream typed comparison](https://github.com/GabrielMarquezMatte/ExcelReader/blob/ca5b50f99e8ef57ab476f0a2bc8043558d58b28d/tests/ExcelReader.Benchmarks/Parse/ParseBenchmark.cs).
These typed methods do not compare UTF-8 spans with materialized strings.

Run validation before measurements, from the repository root:

```powershell
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --validate-bdn
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --validate
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --filter '*TypedReadBenchmarks*' --job Dry --artifacts ./Ignore/Benchmarks/typed-reader-dry --noOverwrite
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --filter '*TypedReadBenchmarks*' --warmupCount 8 --iterationCount 10 --launchCount 1 --outliers DontRemove --artifacts ./Ignore/Benchmarks/typed-reader --noOverwrite
```

`--validate-bdn` invokes the pinned harness's compilation validator for every
enabled workload and parameter case. It checks declaration compatibility before
fixture qualification; a Dry run also proves generated-process execution.

Source is the default OfficeIMO input. For a published-package comparison,
pass an explicit version when building or running the project:

```powershell
dotnet run -c Release -p:OfficeIMOBenchmarkPackageVersion=3.4.4 --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --filter '*TypedReadBenchmarks*' --job Dry --artifacts ./Ignore/Benchmarks/typed-reader-package-dry --noOverwrite
```

For a before/after comparison using saved builds, pass
`-p:OfficeIMOBenchmarkAssemblyDirectory=<absolute-directory>`. That directory must
contain `OfficeIMO.Excel.dll` and its matching `OfficeIMO.Core.dll`. The runner
prints the selected directory, assembly versions, and SHA-256 hashes. Both the
saved-assembly selection and the chosen package version reach BenchmarkDotNet's
generated build. Keep source, saved-build, and package reports separate and record
the source commits. A package version and saved directory cannot be combined.

Additional sizes and coordinate variants are explicit choices:

```powershell
$env:OFFICEIMO_TYPED_BENCHMARK_ROWS = '1000,50000,250000'
$env:OFFICEIMO_TYPED_BENCHMARK_SHAPES = 'Original,SharedStrings,RowReferences,Dimension'
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --validate
```

`RowReferences` adds sequential row indices. `Dimension` adds `A1:D{rows + 1}`.
Both preserve all values, styles, and omitted cell references. The same generated
bytes go to every reader in each case. These shapes diagnose metadata sensitivity;
they do not replace the original producer's input. Each size/shape adds three
typed benchmark cases.

`SharedStrings` uses the upstream `UseSharedStrings = true` producer option with
the same records and absent coordinates/dimension. Setup checks the table's exact
distinct strings and every worksheet text cell's shared-string representation.
The `Shape` column keeps shared-string and inline-string results separate.

`ReaderOpeningBenchmarks` measures only OfficeIMO opening, `FieldCount`, and
disposal. Select it with `--filter '*ReaderOpeningBenchmarks*'` and interpret it
separately: the libraries do different amounts of eager worksheet validation.
Do not subtract results from separate runs to assign an exact fraction of total
time to opening.

Use a topology-derived mask and the existing
`OFFICEIMO_BENCHMARK_PROCESS_PRIORITY` setting when controlling host placement.
Set `OFFICEIMO_BENCHMARK_AFFINITY_MASK` to one nonzero decimal or hexadecimal
mask, such as `0xFFFF0000`. The runner passes this pointer-sized value through
BDN's job API, including masks beyond its command-line option's signed 32-bit
range. BDN's `--affinity` is also available for smaller masks; choose one method
per run.
Apply the same settings to every engine and measure each intended cache domain
separately. Record CPU topology, OS/runtime, power plan, affinity, priority, source
commit, and pinned package versions with results. Retain outliers and inspect
variance and host activity before claiming a small win. The memory column reports
warmed managed allocation, rather than retained or peak process memory.

The exact ExcelReader.NET 6.0.0 package identifies source commit
`ca5b50f99e8ef57ab476f0a2bc8043558d58b28d` and declares the MIT license. The workload
adaptation is attributed in [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
Sylvan.Data 0.2.17 and Sylvan.Data.Excel 0.5.8 are pinned for this lane. Check the
complete BDN report for build or validation failures; process success alone does
not establish that all expected benchmarks ran.

## Raw reads

`RawReadBenchmarks` uses the upstream no-header XLSX/XLSB writer workload and
BIFF8/OLE XLS generator, with the same 50,000 records and values. Setup validates
every field, row order, numeric/date type, row count, single-sheet count, and
checksum. OfficeIMO calls `GetValue` and switches on the object type; ExcelReader
reads string byte spans and parses numeric/date cells; Sylvan dispatches on cell
types and formats. The timed loops use the upstream public APIs and include a
count/checksum check. Their string and boxing work differs, so this original
API comparison is labeled separately.

`MaterializedRawReadBenchmarks` changes ExcelReader's string access to `GetString`
and keeps the other operations the same. It provides the upstream materialized
XLSX comparison and the corresponding XLSB/XLS cases. All readers now materialize
strings, while OfficeIMO's `GetValue` still exposes boxed values. Keep allocation
and API-work differences visible when interpreting this comparison.

The qualified Sylvan raw lane corrects two upstream option/dispatch mismatches.
Its default schema consumes the first data row as a header, returning 49,999 rows
starting at ID 2. The lane uses `ExcelSchema.NoHeaders` so all 50,000 rows are read.
Styled numeric dates report `ExcelDataType.Numeric`; the lane checks
`GetFormat(column).Kind` and reads date ticks for date-formatted cells. Otherwise
the accumulator would consume an OA serial value instead of the same date ticks.
Setup prints the default-header diagnostic separately. Typed scans have a real
header and use explicit date getters, so these corrections apply only to raw reads.

```powershell
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --validate-raw
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --filter '*RawReadBenchmarks*' --job Dry --artifacts ./Ignore/Benchmarks/raw-reader-dry --noOverwrite
```

The three format values add nine cases per class. The existing
`OFFICEIMO_TYPED_BENCHMARK_ROWS` setting also selects raw input sizes; the upstream
XLS fixture supports at most 65,536 rows. Coordinate variants apply only to typed
XLSX. Source, saved-build, and package switches work for both classes.
Small XLS diagnostic inputs pad the Workbook stream to the OLE 4 KiB regular-stream
cutoff. This fixes the upstream generator's missing mini stream for small inputs
and does not alter the original 50,000-row fixture.

For CPU sampling of the exact OfficeIMO typed scan, run the compiled runner with
`--profile-officeimo 500`. It generates and validates the input once, performs 50
complete scans to warm up, then consumes 500 scans and reports an aggregate
observation. This diagnostic loop emits no performance table; use BenchmarkDotNet
for timing and allocation comparisons.

| Comparison dependency | Version | License in the exact package |
| --- | --- | --- |
| BenchmarkDotNet | 0.15.8 | MIT |
| ExcelReader.NET | 6.0.0 | MIT |
| Sylvan.Data | 0.2.17 | MIT (`license.txt`) |
| Sylvan.Data.Excel | 0.5.8 | MIT (`license.txt`) |
| System.IO.Hashing (setup integrity validation) | 10.0.12 | MIT |

## XLSX writes

`WriterBenchmarks` uses the prepared `List<TypedRecord>` and the two default writer
operations from the [upstream writer comparison](https://github.com/GabrielMarquezMatte/ExcelReader/blob/ca5b50f99e8ef57ab476f0a2bc8043558d58b28d/tests/ExcelReader.Benchmarks/Write/WriteBenchmark.cs).
Each timed invocation creates a `MemoryStream` with a 4 MiB capacity, writes the
four-column header and all prepared records, fully finalizes the package, and
returns its length. Record preparation is outside measurements. OfficeIMO calls
its ordinary `ExcelDocument.WriteRows` API with default options; ExcelReader calls
its default `XlsxWorkbookWriter` API and awaits finalization.

The defaults produce different package sizes and metadata. OfficeIMO writes
explicit row/cell references, a dimension, and document properties; ExcelReader
omits the coordinates and dimension. These remain part of the measured operations.
Setup validates every header and field, date styles and date system, sheet count
and name, coordinate policy, ZIP entry lengths and CRCs, well-formed XML, and
relationship targets. Both outputs are then read completely by ExcelReader and
Sylvan, checking every value and row order. Output validation is outside timings.

```powershell
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --validate-write
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --filter '*WriterBenchmarks*' --job Dry --artifacts ./Ignore/Benchmarks/writer-dry --noOverwrite
```

The row-count setting and source, saved-build, and package switches apply to
writes too. For CPU sampling of the same OfficeIMO operation, the compiled runner
accepts `--profile-officeimo-write 500`: setup runs once, 50 writes warm up the
process, and 500 writes contribute to an aggregate package length. This diagnostic
loop reports no timing or allocation measurements.

## Async enumeration

`TypedAsyncReadBenchmarks` uses OfficeIMO.Core's automatic `RowsAsAsync<TypedRecord>`
projection, ExcelReader's async workbook creation and parser enumeration, and
Sylvan's `CreateAsync` plus `GetRecordsAsync<TypedRecord>`. Each returns the same
materialized model, and setup validates every field in every async result.
Original and shared-string inputs use the same `Shape` parameter as synchronous
typed scans. OfficeIMO's synchronous original typed benchmark keeps its explicit
typed-getter loop; the async automatic mapping work is measured separately.

`TypedXlsbReadBenchmarks` and `TypedXlsbAsyncReadBenchmarks` generate the upstream
header-bearing XLSB input using `XlsbSheetWriter.WriteRecordsAsync` with default
options. Each class materializes and consumes the same four-field model; setup
checks every field, header, row order/count, and single-sheet count. Select these
classes separately from XLSX input shapes.

`RawAsyncReadBenchmarks` reads the same no-header XLSX/XLSB/XLS inputs using each
library's existing public async methods. `MaterializedRawAsyncReadBenchmarks`
additionally materializes ExcelReader's string cells, preserving the same work
distinction as synchronous raw scans. Setup checks async field values, types,
order, and counts before any measurements.

The inputs are memory streams. OfficeIMO opens these bytes synchronously; its
`ReadAsync` currently executes parsing synchronously and returns a completed task.
The lane exercises the public async contract without adding a thread-pool wrapper.
It measures enumeration and API overhead on memory input and does not establish
asynchronous disk/network I/O behavior or scalability under blocked I/O.

```powershell
$env:OFFICEIMO_TYPED_BENCHMARK_SHAPES = 'Original,SharedStrings'
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --validate-async
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --filter '*TypedAsyncReadBenchmarks*' '*RawAsyncReadBenchmarks*' --job Dry --artifacts ./Ignore/Benchmarks/async-reader-dry --noOverwrite
```

## Real-data and string-heavy inputs

`RealDataReadBenchmarks` uses the exact `65K_Records_Data` XLSX, XLSM, XLSB, and XLS
files from the [pinned upstream real-data suite](https://github.com/GabrielMarquezMatte/ExcelReader/blob/ca5b50f99e8ef57ab476f0a2bc8043558d58b28d/tests/ExcelReader.Benchmarks/RealData/RealDataReadBenchmark.cs).
Each has 65,535 data rows, one header, and fourteen columns. The XLSX/XLSB/XLS bytes
match the existing shared `MarkPflug65KFixture` hashes; XLSM is pinned to SHA-256
`6E4B3A60D3C59C075BC53E4AF6DAE482DEC8532A0AF0309DA4EFB6AD59F39ABE`.
The shared fixture loader downloads missing files and verifies their hashes in
setup. Set `OFFICEIMO_BENCHMARK_DATA` to a fixture directory to reuse local copies.
The project does not redistribute the data files.

Timed real-data scans read prepared bytes through `MemoryStream`, matching the
upstream methods. Explicit memory cases use ExcelReader's `ReadOnlyMemory<byte>`
API and OfficeIMO's `byte[]` API. OfficeIMO's `Stream` case uses its normal stream
overload, including its input-buffering cost. None of these methods measures disk
I/O. `PrefetchedRealDataReadBenchmarks` preserves the upstream ZIP decompression
prefetch cases for XLSX/XLSM/XLSB; `MaterializedRealDataReadBenchmarks` reads
materialized strings in every engine.

`StringHeavyReadBenchmarks` uses the [upstream generator](https://github.com/GabrielMarquezMatte/ExcelReader/blob/ca5b50f99e8ef57ab476f0a2bc8043558d58b28d/tests/ExcelReader.Benchmarks/Shared/StringHeavyWorkbookGenerator.cs)
with shared strings enabled: 65,536 data rows, one header, eight text columns, and
`Id`, `Amount`, and `Date`. It preserves the original deterministic text pools
and distributions. `MaterializedStringHeavyReadBenchmarks` includes ordinary and
explicitly interned ExcelReader string access as separate methods. OfficeIMO and
Sylvan use their normal reader defaults; cache settings are not silently matched
through adapters. `OFFICEIMO_STRING_BENCHMARK_ROWS` selects diagnostic sizes.

Setup compares every field, native field type, row order, count, and width across
all readers; generated string-heavy values are also compared with the original
generator's formulas. Each timed option is exercised once in setup. All readers
include the header row using `HasHeaderRow = false` or `ExcelSchema.NoHeaders`.
This corrects the upstream Sylvan header-skipping and numeric-date dispatch
differences so each method consumes the same cells. Span and materialized-string
checksums are tracked separately, including their UTF-8 byte versus character
length distinction.

```powershell
$env:OFFICEIMO_BENCHMARK_DATA = './Ignore/Benchmarks/fixtures/65k'
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --validate-real
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --validate-strings
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --filter '*RealDataReadBenchmarks*' '*StringHeavyReadBenchmarks*' --job Dry --artifacts ./Ignore/Benchmarks/large-input-dry --noOverwrite
```

## Writer options and native binary formats

`SharedStringWriterBenchmarks` selects shared strings in both ordinary XLSX
writers. `CompactWriterBenchmarks` uses the public `IncludeCellReferences = false`
option and retains the original inline-string peer writer and its write-prefetch
variant. The ordinary `WriterBenchmarks` case continues to use default options.
OfficeIMO's coordinate option omits data-row coordinates; the header retains its
one row and four cell references. Setup checks the actual selected string and
coordinate policies, every field, date styles, package CRCs, and relationships.

`NativeWriterBenchmarks` compares XLSB using OfficeIMO's normal `Create`, `AddWorksheet`,
`InsertObjects`, and native `Save` APIs. Model construction and saving are both
inside the measured operation. The list of source records is prepared outside
measurement for both libraries. XLSB preserves the upstream four-cell row buffer
and 4 MiB output capacity; XLS preserves its separate 16 MiB capacity. Both
operations finish and dispose the workbook before returning the output length.
`ConfiguredXlsbWriterBenchmarks` preserves the peer's shared-string and write-
prefetch variants without inventing equivalent OfficeIMO options.

Every native output is read field by field through ExcelReader and Sylvan outside
timing. XLSB also checks ZIP payload CRCs, required parts, and relationships.
`NativeXlsWriterDiagnosticBenchmarks` preserves the original manual XLS methods
with 16 MiB output capacity; `XlsRecordWriterDiagnosticBenchmarks` preserves the
typed methods' 4 MiB capacity. These classes have no paired baseline ratio.
The XLS qualification logs `INDEX` presence and its `DBCell` offsets separately
from readable cell values. [MS-XLS worksheet grammar](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/f41c06f2-9057-49a1-8c3f-a4a4d211fc56)
requires an `Index` record; its [cell-table grammar](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/a1b3d8b4-7442-41fd-9c57-bbd2a6394082)
requires `DBCell` records. Independent sequential readability alone establishes
the cells returned by those readers; it does not establish full Excel format
conformance. Interpret XLS timing alongside this structure observation. No Excel
repair or data-loss behavior is inferred from it. ExcelReader.NET 6.0.0's generated
XLS output omits the `INDEX` record and is rejected by Excel 16.0 build 20430 in
normal load mode; the OfficeIMO output and pinned real-data controls load normally.
These XLS methods remain visible as diagnostics and cannot support an equivalent-
valid-output speed ranking. This is separate from their successful sequential
cell-value checks through ExcelReader and Sylvan.

```powershell
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --validate-write-options
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --filter '*SharedStringWriterBenchmarks*' '*CompactWriterBenchmarks*' '*NativeWriterBenchmarks*' '*ConfiguredXlsbWriterBenchmarks*' --job Dry --artifacts ./Ignore/Benchmarks/writer-options-dry --noOverwrite
```

Set `-p:OfficeIMOBenchmarkNewApis=true` when compiling optional cases that call
current source APIs. The selection is recorded and passed to BenchmarkDotNet's
generated build. Leave it disabled when qualifying older package or assembly
selectors. An older writer that ignores an explicitly requested shared-string
policy fails that case's setup rather than producing a misleading comparison.

`RecordWriterBenchmarks` follows the upstream attribute-layout `WriteRecordsAsync`
methods for XLSX/XLSB/XLS. `MappedRecordWriterBenchmarks` preserves the explicit
`IExcelRecordMap` layout for XLSX/XLSB; the upstream suite has no mapped XLS method.
All record-writing cases use a newly allocated 4 MiB output stream, including XLS,
distinct from the original manual XLS writer's 16 MiB capacity. OfficeIMO uses
its normal callback writer for XLSX and model construction plus native save for
binary formats. Setup compares every output field through independent readers.
The typed XLS output shares the manual peer writer's Excel normal-load rejection
and remains in its separate diagnostic class.

`Utf8WriterBenchmarks`, enabled with `OfficeIMOBenchmarkNewApis=true`, prepares
the upstream 100,000 by four `customer-{column}-{row % 5000:D5}` UTF-8 string input
as unmanaged offset/data buffers. Each invocation resets the same 32 MiB output
stream. OfficeIMO consumes these spans through normal `WriteRows` and `WriteUtf8`,
including row enumeration and callback costs. The package comparison calls the
public `XlsxRowWriter.WriteUtf8` API. The upstream `NativeStrings_Xlsx` method
additionally dispatches through its internal, source-only `NativeColumn` and
`WriteApi`; this public companion is labeled separately and does not stand in
for that original internal method. This case writes no header and qualifies all
400,000 field values, native types, row order/count, ZIP CRCs, and relationships.

For native output inspection, set `OFFICEIMO_BENCHMARK_WRITTEN_ARTIFACTS` to an
existing explicit output directory before qualification. Setup retains the
validated XLS/XLSB packages there; ordinary benchmark runs retain only compact
structure and size descriptions.

```powershell
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --validate-write-records
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -p:OfficeIMOBenchmarkNewApis=true -- --validate-write-utf8
```

## Optional CSV and Arrow builds

Set `OfficeIMOBenchmarkCsv=true` to include the CSV classes, Sylvan.Data.Csv 1.4.4,
and the matching OfficeIMO.CSV source, package, or saved assembly. XLSX-only saved
builds continue to require only Excel and Core. The CSV saved lane additionally
requires `OfficeIMO.CSV.dll` in the selected directory. Borrowed UTF-8 API cases
also require `OfficeIMOBenchmarkNewApis=true`; materialized and ordinary typed
CSV cases compile against the older published-package selector without that flag.
The decoded-text parallel date methods also require the current-API flag:
published 3.4.4 has a date-batch defect in that path. Its qualified Core parallel
methods remain available, and qualification prints the known exclusion.
Both flags reach the generated BDN build, and every loaded engine's informational
version and SHA-256 hash are printed and checked against saved inputs.

The CSV suite covers raw, materialized, typed, async, 32-column text, and writing
work. `--validate-csv` checks every field and count outside timing. The parallel
suite compares complete typed records separately from native aggregate work.
`--validate-csv-parallel` uses bounded 50,000-row fixtures across the upstream
degrees of parallelism; measured cases retain their original 4.3M, 8M, and 3M
sizes. `OFFICEIMO_CSV_PARALLEL_BENCHMARK_ROWS` and
`OFFICEIMO_CSV_PARALLEL_BENCHMARK_DOP` select diagnostic sizes and degrees.
Set `OFFICEIMO_BENCHMARK_OUTPUT` to an existing task-owned scratch directory for
parallel CSV input files; cleanup removes the generated input after each case.

`CsvDirectAggregateBenchmarks` calls the public direct reduction APIs with an
independent state per partition or batch and an associative merge. The peer
retains its original `CsvParallel.AggregateAsync` operation. OfficeIMO's text
case reads and decodes the file inside the measured operation, then calls
`AggregateTextRowsAsParallel`. Its two asynchronous cases call
`AggregateRowsAsParallelAsync` with the file path or an asynchronous `FileStream`;
both include input opening, actual asynchronous reads, decoding, field parsing,
bounded worker batches, merging, and disposal. No input text is preloaded for
those asynchronous methods.

The output contract parses all three integer fields, or both text fields, the
date, both decimals, and the units field before reduction. The original aggregate
is the sum of `A` or `Units`, together with the complete row count. Qualification
checks every parsed field and each generated row pattern's exact multiplicity,
including repeated heavy records, rather than relying on the final sum alone.
A qualification-only stream probe confirms asynchronous reads without adding a
wrapper to measured input. OfficeIMO's bounded asynchronous parser materializes
decoded UTF-16 fields; the peer parses UTF-8 fields. These implementation and
allocation policies remain visible in results.

The default inputs retain 8M integer rows, 4.3M conversion-heavy rows, and the 3M
heavy variant with the configured DOP matrix. The short command uses 50K rows per
shape and degree unless `OFFICEIMO_CSV_PARALLEL_BENCHMARK_ROWS` selects another
size. The older `CsvParallelAggregateBenchmarks` caller loop over mapped results
is a separate `CsvCallerReductionDiagnostic` category without paired ratios.

```powershell
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -p:OfficeIMOBenchmarkCsv=true -p:OfficeIMOBenchmarkNewApis=true -- --validate-csv-direct-aggregate
```

The explicit `--validate-csv-direct-aggregate-scale` command uses the measured
input sizes and the selected DOP values, while retaining the same per-field and
multiplicity proof. Clear `OFFICEIMO_CSV_PARALLEL_BENCHMARK_ROWS` to qualify the
original 4.3M, 8M, and 3M inputs. For a bounded choice of degrees:

```powershell
$env:OFFICEIMO_CSV_PARALLEL_BENCHMARK_DOP = '1,4'
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -p:OfficeIMOBenchmarkCsv=true -p:OfficeIMOBenchmarkNewApis=true -- --validate-csv-direct-aggregate-scale
```

`CsvRealDataReadBenchmarks` preserves the stream and direct-memory reads of the
pinned fourteen-column `65K_Records_Data.csv`, including its header. The shared
fixture hash is
`AC959F43CF1077B71D310E6E49E3C168BA63A448F1855D45F44E734273EBA490`.
The peer and current OfficeIMO borrowed APIs expose UTF-8 bytes; Sylvan's
`GetFieldSpan` exposes UTF-16 characters. Setup prints both totals and validates
every string and borrowed span through both peer input routes. These original
span results retain their encoding distinction. The separate
`CsvRealDataMaterializedReadBenchmarks` methods all consume materialized strings.
All inputs come from prepared bytes; these cases do not measure disk reads.

`CsvRecordWriterBenchmarks` retains the original 50,000-record attribute and
generated-layout CSV writers in separate `Mapped` cases. Both engines receive
prepared lists and a new 4 MiB destination. OfficeIMO invokes ordinary
`WriteObjects` object mapping for each model; it does not have the peer's
generated-layout API. The operation includes mapping, serialization and complete
finalization. Setup checks every header, field and row order with OfficeIMO and
Sylvan. The existing explicit row-writing comparison remains separate.

```powershell
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -p:OfficeIMOBenchmarkCsv=true -p:OfficeIMOBenchmarkNewApis=true -- --validate-real-csv
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -p:OfficeIMOBenchmarkCsv=true -- --validate-csv-record-write
```

`OfficeIMOBenchmarkArrow=true` includes the optional Arrow owner and
ExcelReader.Arrow 6.0.0. It requires both CSV and current-API flags. The saved
assembly lane additionally requires `OfficeIMO.Data.Arrow.dll`; Apache.Arrow
23.0.0 supplies the matching binary contract. The default build loads none of
these optional Arrow engines. Arrow schema inference policies must be qualified
separately when field types or nullability differ.

`ArrowConversionBenchmarks` uses the exact 100,000-record, one-batch explicit
schema cases for typed CSV, typed XLSB, and eight-string CSV. Bounded OfficeIMO
cases use 8,192-record batches, producing thirteen batches, and remain separate
from the one-batch results. Setup checks every field value, schema name/type,
required-field nullability, timestamp units, batch count, and complete row count.
Inference is a separate diagnostic without a paired ratio: ExcelReader returns
four required string fields, while OfficeIMO returns string, Int32, microsecond
timestamp, and Decimal128(29,10) fields under its ordinary inference policy.

Sylvan.Data.Csv 1.4.4 and ExcelReader.Arrow 6.0.0 use MIT; Apache.Arrow 23.0.0 and
its Apache.Arrow.Scalars dependency use Apache-2.0. These dependencies remain in
this opt-in comparison project.

```powershell
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -p:OfficeIMOBenchmarkCsv=true -p:OfficeIMOBenchmarkNewApis=true -- --validate-csv
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -p:OfficeIMOBenchmarkCsv=true -p:OfficeIMOBenchmarkNewApis=true -- --validate-csv-parallel
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -p:OfficeIMOBenchmarkCsv=true -p:OfficeIMOBenchmarkNewApis=true -p:OfficeIMOBenchmarkArrow=true -- --validate-arrow
```

## ADO access and first use

`AdoReadBenchmarks` opens the original typed XLSX bytes through `IDataReader`
interface dispatch and consumes every field with `GetValue` or typed getters.
The current-API build also exposes the UTF-8 text-copy case using the public
bounded-copy contract. Setup checks the returned fields, types, order, and count
for each accessor, including the actual copied bytes and count for every text
field in the UTF-8 case. `DataTableLoadDiagnostics` keeps ordinary `DataTable.Load`
visible separately: the reader bridges produce different stored column types,
including culture-formatted date text. It validates those schemas and all
normalized values, prints the schema difference, and has no paired baseline ratio.

`ColdStartReadBenchmarks` and `ColdStartWriteBenchmarks` measure one first mapping
or automatic-layout use in each of sixteen fresh processes. Setup builds input
through low-level cell writes and prepares records without touching either
library's typed mapping cache. The separate `--validate-cold` command checks
reader fields and complete writer outputs in its own qualification process;
running it does not warm a later benchmark child. The timing measures first use.
BenchmarkDotNet 0.15.8's memory diagnoser obtains allocation statistics from an
additional invocation in the same child after mapping/layout initialization
([engine source](https://github.com/dotnet/BenchmarkDotNet/blob/v0.15.8/src/BenchmarkDotNet/Engines/Engine.cs#L126)).
Label this class's `Allocated` column as second-use allocation; it does not
measure cold initialization allocation. Keep these timing and allocation scopes
distinct when publishing results.

```powershell
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --validate-ado
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --validate-cold
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --filter '*ColdStartReadBenchmarks*' '*ColdStartWriteBenchmarks*' --artifacts ./Ignore/Benchmarks/first-use --noOverwrite
```

## Model strategies and shared-string lifecycle

`TypedModelReadBenchmarks` compares public automatic mapping to the same class
and struct property shapes on the original 50,000-row input. The original manual
OfficeIMO method remains in `TypedReadBenchmarks`; automatic mapping gets its
own results. `BorrowedRefStructReadBenchmarks` requires the current-API flag and
consumes a borrowed UTF-8 name plus the same three typed fields. ExcelReader maps
that ref struct automatically; OfficeIMO uses `RowsAsBorrowed<T>` with a caller
factory over public borrowed-span and typed-getter APIs. The manual-loop method
is a separate diagnostic control. The method names state the mapping strategy;
these results do not establish an automatic ref-struct property-mapping API.
All model cases qualify every field, count, and row order before measurement.

`RealDataTypedReadBenchmarks` maps the fourteen-property class from the pinned
65K XLSX/XLSB files, including both date fields, Int64 identifiers and units, and
all five Double monetary fields. Both readers use their ordinary automatic
header mapping. Sylvan supplies an independent per-field oracle outside timing.

`SharedStringFirstRowBenchmarks` measures initialization, moving to the first
header row, and disposal. `SharedStringFullScanBenchmarks` consumes all
materialized fields separately. `SharedStringBorrowedScanBenchmarks` retains
the original peer span reads and its explicit prefetch option as diagnostics.
`SharedStringUtf8ReadBenchmarks` adds the public OfficeIMO borrowed shared-string
companion in a current-API build. It uses the same generator shape with an
explicit 100,000-row default and a `Format` parameter for native XLSX and XLSB;
the original XLSX lifecycle classes retain 65,536 rows. Each format uses the same
eight text fields, Id, Amount, Date, and header, with deflated and stored ZIP
variants. Native XLSB calls the public reader's own borrowed UTF-8 provider;
there is no caller string-to-byte adapter.
`OFFICEIMO_SHARED_UTF8_BENCHMARK_ROWS` selects a labelled diagnostic size.
Its text fields use public borrowed UTF-8 spans, and numeric/date values retain
their native getters. Setup validates every borrowed byte, native field type,
count, and order against the generator, peer, and independent Sylvan reader.
Each measured operation opens and disposes a new reader. OfficeIMO normalizes
shared-string text and encodes UTF-8 on lookup, with a bounded cache owned by that
reader. Native XLSB decodes its UTF-16 string storage and uses its canonical
bounded UTF-8 cache. First lookup, eviction, and uncached encoding costs remain
timed in both formats. The
peer's normal and explicit decompression-prefetch policies remain separate.
Borrowing a returned span does not imply allocation-free shared-string lookup.
The materialized class provides the separate string-producing comparison.
The original lifecycle classes use the 65,536-row string-heavy
fixture and a stored ZIP variant. The pinned upstream rewrite stores every ZIP
part without compression; qualification checks identical inflated part hashes,
ZIP CRCs, and every workbook field. Initialization policies differ between
libraries, so first-row results are lifecycle costs. Do not subtract independent
means and present the result as a measured worksheet-only stage.

The upstream `IndexParse_Utf8Parser` and `IndexParse_DigitLoop` methods compare
standalone integer parsers over pre-extracted SST index buffers. They do not call
a workbook reader and have no OfficeIMO workbook counterpart.

`StyledRowWriterBenchmarks` preserves the upstream `StyledRows_Xlsx` public
operation: 100,000 numeric values from 0 through 99,999, one cell per row, no
header, a named bold row default, and a retained 32 MiB output stream. OfficeIMO
uses ordinary `WriteRows` with a declared `ExcelStyleDefinition` catalog and
`DefaultRowStyle`; enumeration, callbacks, style compilation, package creation,
and full finalization remain in the measured operation. Ordinary OfficeIMO row
and cell coordinates stay enabled. Both writers author resolved bold cell XFs
as well as real row defaults; that work makes the existing values display bold
in Excel, rather than only formatting subsequently entered cells.

Qualification checks all 100,000 numeric values and native types with all three
readers, every row/cell XF and its font, stylesheet counts and references,
coordinate order, ZIP CRCs, and package relationships. The current-API command
is `--validate-write-styled-rows`. Set `OFFICEIMO_STYLED_WRITE_ARTIFACTS` to a new,
existing absolute directory to preserve the two qualified workbooks for native
Excel inspection; normal BDN setup keeps outputs in memory.

```powershell
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -p:OfficeIMOBenchmarkNewApis=true -- --validate-models
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --validate-real-typed
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -- --validate-shared-stages
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -p:OfficeIMOBenchmarkNewApis=true -- --validate-shared-utf8
```

The shared UTF-8 qualification command defaults to the complete 100,000 data
rows in both formats and ZIP variants. A separate reduced-row Dry run checks
that BenchmarkDotNet can build and execute all twelve generated workload cases:

```powershell
$env:OFFICEIMO_SHARED_UTF8_BENCHMARK_ROWS = '100'
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -p:OfficeIMOBenchmarkNewApis=true -- --filter '*SharedStringUtf8ReadBenchmarks*' --job Dry --artifacts ./Ignore/Benchmarks/shared-utf8-dry --noOverwrite
Remove-Item Env:\OFFICEIMO_SHARED_UTF8_BENCHMARK_ROWS
```

Keep the 100K contract proof and 100-row generated-process check distinct. Dry
means and allocations do not establish a performance ranking.

`ArrowStringWriterBenchmarks` uses the exact prepared 100,000-row batch with four
required string fields and a retained 32 MiB output stream. ExcelReader invokes
its public `WriteRecordBatch` method. OfficeIMO caller code reads the same Arrow
UTF-8 spans and passes them through ordinary `WriteRows`/`WriteUtf8`, including
enumeration, callback, column access, headers, and complete finalization in the
measured operation. This benchmark-only orchestration supports the fixture's
four string columns; it is not a production Arrow-to-XLSX adapter. Qualification
checks all headers and 400,000 values, types, order/count, ZIP CRCs, and
relationships. The ordinary OfficeIMO coordinate defaults remain enabled.

```powershell
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -p:OfficeIMOBenchmarkCsv=true -p:OfficeIMOBenchmarkNewApis=true -p:OfficeIMOBenchmarkArrow=true -- --validate-write-arrow-strings
```

## Authenticated encrypted workbooks

The current-API build adds the exact small ciphertext and generated large
workload from the pinned
[encrypted benchmark](https://github.com/GabrielMarquezMatte/ExcelReader/blob/ca5b50f99e8ef57ab476f0a2bc8043558d58b28d/tests/ExcelReader.Benchmarks/Crypto/EncryptedWorkbookBenchmark.cs).
The original small XLSX is 15,360 bytes, SHA-256
`13F9532B916634CB112D3593E7AB859F67B2DA35A3A47E574F1C1E26C5DA370D`.
Its 8,915-byte paired plaintext is an independent decryption oracle, rather
than a replacement encrypted input. Fixture provenance, the original MIT
license, plaintext hash, and independent authentication evidence are documented
in `OfficeIMO.TestAssets/Documents/ExcelEncryptionCorpus/SOURCE.md`. The original
encryption producer is unpinned. The public fixture password is `hunter2`.

The generated large input defaults to 300,000 rows with no header and the same
four fields as the original raw fixture. Preparation uses the pinned workbook
generator shape and public `Excel.EncryptPackage`, with Agile AES-256-CBC,
SHA-512, 100,000 password iterations, and a complete package integrity HMAC.
`EXCELREADER_LARGE_ENCRYPTED_ROWS` selects an explicitly different row count.
Preparation authenticates both readers and compares every field and native type
against a Sylvan scan of the plaintext and the generator's expected values.
It writes immutable plaintext/ciphertext files and a manifest recording their
lengths, SHA-256 hashes, producer commit, and encryption settings. Measured setup
requires this existing cache and rejects changed bytes; it never re-encrypts it.

`EncryptedVerifiedReadBenchmarks` opens, authenticates the full ciphertext,
consumes every materialized text/numeric/date field, and disposes the reader.
Both implementations perform full integrity verification. ExcelReader uses
`GetString`, numeric parsing, and date getters; OfficeIMO uses ordinary
`GetValue`, including its boxed native values. The `MemoryInput` column keeps
byte/memory input separate from the original stream type: `FileStream` for the
small fixture and `MemoryStream` for the generated input. Qualification checks
all sync and async access paths field by field, in addition to the complete
count/checksum. OfficeIMO eagerly buffers both ciphertext and decrypted package;
the peer stream reader decrypts lazily after its required integrity pass.
Allocation and lifecycle results include those implementation policies.

`EncryptedOriginalSmallDiagnostics` and `EncryptedOriginalLargeDiagnostics`
retain the upstream row-width-only and open-only operations in separate tables
without paired ratios. These methods do not consume field values. The original
stream methods disable integrity except for the explicitly named verify method;
the public memory opener always authenticates even with the default option.
Their costs therefore cannot rank an authenticated full-field scan.
`EncryptedVerifiedAsyncReadBenchmarks` uses both public async stream openers and
async enumeration for the authenticated full-field contract. Cryptographic work
remains CPU work; an async method does not imply parallel decryption.

Set existing absolute corpus and task-cache directories, then prepare once and
qualify before measuring. Use the same cached files for before/after runs:

```powershell
$env:OFFICEIMO_ENCRYPTED_BENCHMARK_CORPUS = (Resolve-Path ./OfficeIMO.TestAssets/Documents/ExcelEncryptionCorpus).Path
$env:OFFICEIMO_ENCRYPTED_BENCHMARK_CACHE = '<existing-absolute-task-cache-directory>'
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -p:OfficeIMOBenchmarkNewApis=true -- --prepare-encrypted
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -p:OfficeIMOBenchmarkNewApis=true -- --validate-encrypted
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -p:OfficeIMOBenchmarkNewApis=true -- --filter '*EncryptedVerifiedReadBenchmarks*' --artifacts ./Ignore/Benchmarks/encrypted-verified --noOverwrite
```

`TypedStreamAsyncReadBenchmarks` adds public asynchronous stream opening to the
existing typed XLSX async comparison. Both peers already open streams
asynchronously. OfficeIMO now uses `OpenDataReaderAsync(Stream)` and the ordinary
async typed projection. The original byte-open OfficeIMO enumeration lane stays
in `TypedAsyncReadBenchmarks`. The stream-opening companion includes input
buffering, discovery, mapping, and disposal, and qualifies every record field.
Run it with the current-API flag and `--validate-stream-async`.

`CsvUtf8WriterBenchmarks` uses the same prepared 100,000-by-four UTF-8 offset/data
input and retained 32 MiB destination as the native CSV workload. The measured
OfficeIMO caller uses the public streaming `WriteUtf8Row` contract; the peer uses
its public UTF-8 row writer. The original internal native write API remains a
separate source diagnostic. Qualification checks byte-identical output and every
field with independent readers, plus escaping, null/empty, and Unicode cases.

```powershell
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -p:OfficeIMOBenchmarkCsv=true -p:OfficeIMOBenchmarkNewApis=true -- --validate-csv-utf8-writer
```

`ConfiguredXlsbWriterBenchmarks` retains the normal OfficeIMO model/save operation
and adds explicit inline and shared-string save policies in the current-API
build, using `ExcelSaveOptions.XlsbUseSharedStrings = false` and `true`. Prepared
records and the 4 MiB destination capacity remain identical. Model construction,
deduplication when selected, native package creation, and finalization are timed.
The peer's original shared-string and write-prefetch methods remain labelled
separately. Compare shared-string methods with shared-string methods; the default
method's actual storage policy is printed by qualification.

Setup checks every field, native type, date style, header and row order with both
independent readers, together with ZIP integrity and relationships. It also
inspects the actual worksheet string records and SST counts using
[MS-XLSB record framing](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/7bf1de78-9cda-4002-8411-086f79cd4b60)
and the
[record enumeration](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/30e2eaae-0b1e-4e8e-a465-e1ce5575868d).
An inline policy must contain all text in cell records; a shared-string policy
must reference the deduplicated table for every text cell. An unused empty table
does not invalidate otherwise inline output. The short qualification command is:

```powershell
dotnet run -c Release --project ./Benchmarks/ExcelReaderTyped/ExcelReaderTyped.Benchmarks.csproj -p:OfficeIMOBenchmarkNewApis=true -- --validate-write-xlsb-options
```
