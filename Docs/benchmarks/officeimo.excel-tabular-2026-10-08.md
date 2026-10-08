# Excel and CSV tabular benchmarks — 2026-10-08

The published OfficeIMO.Excel 3.4.4 package reproduces the large typed-read gap reported by the pinned upstream suite on this host. The measured OfficeIMO source closes most of that gap on the original four-field workload and allocates about 37% fewer managed bytes than ExcelReader there. The final timing is effectively tied in CPU domain 0 and slower in domain 1. Ordinary automatic mapping, large CSV reductions, Arrow conversion, shared-string initialization and native XLSB writing show important remaining costs. These results support scoped improvements, not a claim that one library wins every workload.

This report contains 652 primary case records: 612 current-source cases, 16 source-start cases, 16 published-package cases and 8 Core-only intervention cases. The 78 supplementary records retain two headline/runtime stages, short and longer native attribution, and recipe probes. Separate historical 130 records comprise 106 generated/real exploratory cases and 24 original headline/Core supporting cases. Failed and eligibility-only runs are retained separately. Counts describe complete selected workload coverage, not statistical certainty or complete coverage of every upstream library.

[Complete primary tables](officeimo.excel-tabular-2026-10-08/tables-primary.md), [identity-preserving primary results](officeimo.excel-tabular-2026-10-08/results-primary.json), [CSV](officeimo.excel-tabular-2026-10-08/results-primary.csv), [supplementary tables](officeimo.excel-tabular-2026-10-08/tables-supplementary.md) and the [raw evidence packet](officeimo.excel-tabular-2026-10-08/README.md) retain every native sample, interval, allocation, outlier, warning and complete case/job/domain identity.

## Original workload and published-package attribution

The pinned [ExcelReader project](https://github.com/GabrielMarquezMatte/ExcelReader/blob/ca5b50f99e8ef57ab476f0a2bc8043558d58b28d/tests/ExcelReader.Benchmarks/ExcelReader.Benchmarks.csproj) uses OfficeIMO.Excel NuGet 3.4.4. Its embedded source is 7ea881d08359b4fddddef8e2aa7782d1ee2265b8. The engineering baseline is source 78bfbd96dc3967b6f3c5e2376ce4c98000213e2c. These are different binaries; their shared three-part version does not prove identity.

The generated XLSX has 50,000 records and one header, with Name/Id/Date/Value, eight repeating ASCII names, Id 1 through 50,000, dates from OADate 45292+(row%3650)+0.25 and Value=row*1.5. The default upstream producer uses inline strings, no dimension and no row/cell coordinates. Setup rejects a different shape and validates every field, order and row count. Timed scans open/dispose the reader, materialize the four-field result and consume all fields.

| Domain | Runtime artifact | OfficeIMO mean ± SD (ms) | ExcelReader mean ± SD (ms) | Office/peer mean | Office bytes/op | Peer bytes/op |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| 0 | Source-start78bfbd96 | 112.129 ± 3.723 | 14.417 ± 0.448 | 7.778 | 2540469 | 4054464 |
| 0 | Published3.4.4 /7ea881d0 | 123.642 ± 3.655 | 12.699 ± 0.435 | 9.736 | 5969752 | 4054464 |
| 0 | Measured source 830951 | 12.095 ± 0.516 | 12.060 ± 0.276 | 1.003 | 2560880 | 4069313 |
| 1 | Source-start78bfbd96 | 107.824 ± 6.915 | 12.126 ± 0.554 | 8.892 | 2621818 | 4069313 |
| 1 | Published3.4.4 /7ea881d0 | 121.815 ± 5.729 | 12.194 ± 1.209 | 9.990 | 6051153 | 4069313 |
| 1 | Measured source 830951 | 10.616 ± 0.841 | 9.659 ± 0.441 | 1.099 | 2560854 | 4069313 |

The published-package rows independently support an approximately 9.7–10.0× mean timing gap on this machine; they do not reproduce the quoted Ryzen 7 5700X timings. The source-start baseline is a separate engineering measurement. Final source 830951 is about 10.2–11.5× faster than the published package by quotients of these sequential recorded means; this is not a paired significance test, a current package-release claim or a universal speedup.

OfficeIMO's original typed method uses ordinary typed getters with caller model construction. ExcelReader uses automatic attribute mapping and Sylvan uses GetRecords. They deliver the same validated fields, with different mapping strategies. The automatic class/struct comparison below prevents a manual construction result from being advertised as automatic-mapper parity.

## Methodology, identities and reliability

The host is Windows 26300.9457 with an AMD Ryzen 9 9950X3D, .NET 10.0.12, SDK 10.0.112 and BenchmarkDotNet 0.15.8. Two recorded affinity domains use masks 65,535 (0xFFFF) and 4,294,901,760 (0xFFFF0000), with High process priority. These are two domains on one host, not independent machines. All outliers remain retained with DontRemove.

| Job family | Launches | Warmups | Measurements | Sizing |
| --- | ---: | ---: | ---: | --- |
| Headline |1|12|15|32 invocations, unroll1|
| Ordinary |1|12|15|1,000ms automatic Pilot target, unroll1|
| Large |1|8|12|1,000ms automatic Pilot target, unroll1|
| Huge CSV |1|4|8|250ms automatic Pilot target, unroll1|
| Core intervention |1|12|15|250ms automatic Pilot target, unroll1|
| Cold |16|0|1 per launch|1 invocation, no Pilot|

All final warm measurements have native WorkloadActual iteration duration at least 100 ms. The declared cold recipe intentionally retains five short-iteration warnings per domain. Native distribution warnings and outliers remain visible; one launch and observed variance still limit small-margin rankings. Allocated means cumulative managed bytes per operation, not retained/peak or unmanaged/native memory.

Actual capture order is retained. Source-start and published headline captures belong to e676, with the original e676-current stage between them. Final 830951 headline captures ran later, domain 0 then domain 1, rather than physically interleaved with the original baseline/package rotation. Benchmark source is byte-identical e676→830951; runtime, capture source, harness, selector and runner identities remain distinct. The 434-case resume uses its own runner/reference-import identities. Exact timestamps and commands are in the [capture ledger](officeimo.excel-tabular-2026-10-08/manifests/capture-ledger.json).

Each group preserves pre/post host snapshots. Counts and persistent CPU deltas cannot establish per-case activity, isolation or causal contamination. In the domain 1 cold capture, persistent testhost process 87244 accumulated 46.578125 CPU seconds during 53.7689242 seconds of capture; persistent processes totaled 63.796875 CPU seconds. This substantial concurrent activity limits cold timing reliability. No automatic deletion/rerun or isolated-host ranking is inferred.

Every actual generated read input has its byte length/SHA256 logged before timing. ZIP inputs additionally log each inflated entry's name/length/SHA256, compressed length and timestamp, except cold input logging hashes only package bytes to avoid warming reader/decompression work. Separately generated packages can have different container hashes despite identical inflated payloads. Retained metadata and payload comparisons do not prove an exclusive cause for every differing container byte. Every implementation inside a read instance consumes the same bytes. Prepared object writer lists are defined and fully field-checked; they have no invented input-file hash. Representative Setup output hashes qualify the output contract, not each timed invocation.

### Measured runtime artifacts

| Component | Source / SHA256 |
| --- | --- |
| Current source |83095146282a96143b23a79a7ecca058a24a709b|
| Unchanged benchmark source |e676775a59458d1095606386c39506ceb8a9bc33|
| OfficeIMO.Excel DLL |F2853A4C62D5C03C5E3F0CBB7CCBB5776FB63CE58CE8933BE4EBAA4F5887502F|
| OfficeIMO.Core DLL |B5565DE58D83BFC91F4FFB69A797CE02DFCB10A5A6409323845D05A5732E1A84|
| OfficeIMO.CSV DLL |B9D19DBE2957F6E2E92CD12D0C8A3003D290621EF2AB3E276953CE5A01BBE4B6|
| OfficeIMO.Data.Arrow DLL |ED4EBF11E9A152C126E28D3BF8FBE588F8431FBAA175920BF8836AB92D5058E8|
| Main/current outer harness |5E79575325D94AE8D09B97CD0C1C341EBB6AC0A24C3EEE6AE4817A897F97E7A4|
| Resume runner |C52377235F70A38C61333009FFA2CC97C0BE2B3C102041CB127587B09AB32A95|
| Generated-build reference hook |CEDF3A831603914A5E024C980CCA65AAAAF25544D98538BC416E963DD8632B12|

Published Excel 3.4.4 DLL SHA256 is BEB5A238B4802DAA4B43244E9E3D69E7349FBA8483D2EF4884D90CC81AAD72D3; Core is 44B3C7657F79296D2F52488328BB14CB7363DDF1784F12E4E8527FD20D4F8745. Their nupkg hashes are respectively 6826377231097A99F01CA3B51C8C7FCFAAF0C37FBBDE7B46FECAAD6A3EC23362 and FA04B42A0DF4CC8F6C7D4FA73C68652B232E02E62A40F824D7F951867A2A534C. Native sidecars retain exact peer/DLL/framework identities too.

## Complete current coverage and interpretation

The selected current matrix has 306 cases per domain. Full native class/parameter/job groups remain authoritative. Do not collapse different classes or parameters into one fastest ratio; some native baseline ratios describe unlike API strategies.

| Selected family | Cases/domain | Work and comparison boundary |
| --- | ---: | --- |
| Original/SST typed headline and ordinary writer |8|Four-field objects; caller mapping and writer metadata policies explicit|
| Generated workbook reads/model kinds/async |54|Raw versus materialized XLSX/XLSB/XLS; automatic class/struct versus borrowed factory; opening versus enumeration|
| Pinned real workbook reads and typed mapping |52|65K/14-column inputs, four workbook formats; stream/bytes and prefetch distinct|
| String-heavy/SST lifecycle and borrowing |47|65,536×8 and100K × 11; deflated/stored, normalization/cache/first-row work distinct|
| ADO getter and explicit UTF8 copying |9|GetValue, typed scalar getters and owned text-byte copying through interfaces|
| Workbook writer paths |26|Ordinary/compact/SST, native model+Save, records, UTF8, Arrow caller rows and true row styles|
| CSV raw/materialized/typed/async/writer/real |37|Three engines; full wide False/True identity and actual output contract|
| Original large CSV DOP1 |27|4.3M / 8M / 3M, mapped objects and direct reductions distinct|
| Original large CSV DOP4 |21|Same full files; three parallel mapping strategies and four direct reductions|
| Explicit/inferred Arrow |7|Six explicit one-owned-batch comparisons; one unpaired peer inference diagnostic|
| Authenticated encrypted workbook scans |12|Original small and generated300K; complete verified plaintext field work, sync/async/actual input route|
| Cold mapping/layout |6|16 fresh processes; first target use after Setup, subsequent diagnostic allocation|

### Generated models and asynchronous reading

| Domain | Class / method | Native parameters | Mean ± SD (ms) | Allocated bytes/op |
| --- | --- | --- | ---: | ---: |
| 0 | TypedModelReadBenchmarks.ExcelReaderAutomatic | RowCount=50000&Model=Class | 10.860 ± 1.548 | 4054600 |
| 0 | TypedModelReadBenchmarks.OfficeIMOAutomatic | RowCount=50000&Model=Class | 17.674 ± 0.335 | 2519646 |
| 0 | TypedModelReadBenchmarks.ExcelReaderAutomatic | RowCount=50000&Model=Struct | 11.787 ± 0.893 | 1660540 |
| 0 | TypedModelReadBenchmarks.OfficeIMOAutomatic | RowCount=50000&Model=Struct | 15.934 ± 0.842 | 176065 |
| 1 | TypedModelReadBenchmarks.ExcelReaderAutomatic | RowCount=50000&Model=Class | 7.528 ± 0.163 | 4054600 |
| 1 | TypedModelReadBenchmarks.OfficeIMOAutomatic | RowCount=50000&Model=Class | 9.978 ± 0.300 | 2479717 |
| 1 | TypedModelReadBenchmarks.ExcelReaderAutomatic | RowCount=50000&Model=Struct | 7.408 ± 0.172 | 1654600 |
| 1 | TypedModelReadBenchmarks.OfficeIMOAutomatic | RowCount=50000&Model=Struct | 10.144 ± 0.261 | 79621 |

Automatic classes and structs use ordinary public mapping. The borrowed ref-struct lane compares ExcelReader automatic mapping with OfficeIMO's public RowsAsBorrowed factory; the caller factory cost is included. It does not establish an arbitrary automatic ref-struct mapper.

Raw original loops expose string materialization, UTF-8 borrowing and scalar-boxing differences; their raw cross-library ratios are API-cost observations. Materialized companions require string production and validate every value/type/order. Async byte-open enumeration, async stream opening and source prefetch remain separate operations. Actual asynchronous input reading does not make synchronous parsing CPU work asynchronous.

### Pinned real data

| Domain | Class / method | Native parameters | Mean ± SD (ms) | Allocated bytes/op |
| --- | --- | --- | ---: | ---: |
| 0 | RealDataTypedReadBenchmarks.ExcelReaderAutomatic | Format=Xlsb | 39.792 ± 0.927 | 8426385 |
| 0 | RealDataTypedReadBenchmarks.OfficeIMOAutomatic | Format=Xlsb | 56.437 ± 1.556 | 30061218 |
| 0 | RealDataTypedReadBenchmarks.ExcelReaderAutomatic | Format=Xlsx | 66.238 ± 2.382 | 8406944 |
| 0 | RealDataTypedReadBenchmarks.OfficeIMOAutomatic | Format=Xlsx | 94.639 ± 1.888 | 9552992 |
| 1 | RealDataTypedReadBenchmarks.ExcelReaderAutomatic | Format=Xlsb | 29.898 ± 1.882 | 8406488 |
| 1 | RealDataTypedReadBenchmarks.OfficeIMOAutomatic | Format=Xlsb | 35.611 ± 1.209 | 25859178 |
| 1 | RealDataTypedReadBenchmarks.ExcelReaderAutomatic | Format=Xlsx | 47.451 ± 2.397 | 8406944 |
| 1 | RealDataTypedReadBenchmarks.OfficeIMOAutomatic | Format=Xlsx | 61.434 ± 3.594 | 8511104 |

The real inputs retain pinned hashes and independent field normalization, including 14 columns. XLSX/XLSM/XLSB/XLS raw access, materialized scans, byte/stream opening and peer prefetch are separate native groups. Ordinary automatic 14-field XLSX/XLSB mapping shows valid throughput/allocation losses; retain them rather than excluding them because the four-field headline is closer.

### Strings and first-row initialization

| Domain | Class / method | Native parameters | Mean ± SD (ms) | Allocated bytes/op |
| --- | --- | --- | ---: | ---: |
| 0 | SharedStringFirstRowBenchmarks.ExcelReaderOpenThroughFirstRow | RowCount=65536&Storage=Deflated | 16.309 ± 0.232 | 908205 |
| 0 | SharedStringFirstRowBenchmarks.OfficeIMOOpenThroughFirstRow | RowCount=65536&Storage=Deflated | 93.223 ± 3.743 | 20955372 |
| 0 | SharedStringFirstRowBenchmarks.SylvanOpenThroughFirstRow | RowCount=65536&Storage=Deflated | 0.100 ± 0.003 | 328832 |
| 0 | SharedStringFirstRowBenchmarks.ExcelReaderOpenThroughFirstRow | RowCount=65536&Storage=Stored | 4.215 ± 0.092 | 801149 |
| 0 | SharedStringFirstRowBenchmarks.OfficeIMOOpenThroughFirstRow | RowCount=65536&Storage=Stored | 71.485 ± 2.003 | 21213609 |
| 0 | SharedStringFirstRowBenchmarks.SylvanOpenThroughFirstRow | RowCount=65536&Storage=Stored | 0.089 ± 0.012 | 327775 |
| 1 | SharedStringFirstRowBenchmarks.ExcelReaderOpenThroughFirstRow | RowCount=65536&Storage=Deflated | 13.268 ± 0.388 | 765240 |
| 1 | SharedStringFirstRowBenchmarks.OfficeIMOOpenThroughFirstRow | RowCount=65536&Storage=Deflated | 68.120 ± 5.528 | 15317486 |
| 1 | SharedStringFirstRowBenchmarks.SylvanOpenThroughFirstRow | RowCount=65536&Storage=Deflated | 0.061 ± 0.004 | 328624 |
| 1 | SharedStringFirstRowBenchmarks.ExcelReaderOpenThroughFirstRow | RowCount=65536&Storage=Stored | 3.229 ± 0.032 | 764160 |
| 1 | SharedStringFirstRowBenchmarks.OfficeIMOOpenThroughFirstRow | RowCount=65536&Storage=Stored | 42.784 ± 3.043 | 15315543 |
| 1 | SharedStringFirstRowBenchmarks.SylvanOpenThroughFirstRow | RowCount=65536&Storage=Stored | 0.058 ± 0.005 | 327769 |

Open-through-first-row is a legitimate caller initialization contract. In this indexed native XLSX path, opening-time worksheet qualification encounters the first shared-string cell reference and loads the complete selected SST, including later unreferenced items, before OpenDataReader returns. The first caller Read selects an already qualified row; this benchmark does not call GetValue or GetString. It is much slower here than peer/Sylvan initialization. This is not replaced with sheet-name discovery, A1 ranges, sampling or unchecked metadata. Full scans measure delivery separately.

Stored/deflated payload comparisons use complete nonempty actual part manifests. UTF-8 borrowing includes normalized text, provider eligibility and cold encoding/cache costs. A derived UTF-8 cache does not mean the original decoded SST has been avoided. XLSB native borrowing is supported; the original 65,536 borrowed-only peer lifecycle remains separate from the 100K public companion.

### ADO access

The nine cases use the same validated 200K field accesses via IDataReader/DbDataReader dispatch. GetValue includes boxed scalar objects; typed getters and explicit owned UTF-8 copies perform different work. Setup verifies copy counts and every text byte, not only returned lengths. GetBytes retains its binary API contract; OfficeIMO GetUtf8Bytes is an explicit text operation. DataTable.Load schema differences remain separate untimed diagnostics and are not advertised as equivalent storage cost.

### Workbook writing and native formats

| Domain | Class / method | Native parameters | Mean ± SD (ms) | Allocated bytes/op |
| --- | --- | --- | ---: | ---: |
| 0 | WriterBenchmarks.ExcelReaderWriter | RowCount=50000 | 10.606 ± 0.289 | 4218069 |
| 0 | WriterBenchmarks.OfficeIMO | RowCount=50000 | 18.648 ± 0.878 | 4238334 |
| 1 | WriterBenchmarks.ExcelReaderWriter | RowCount=50000 | 9.854 ± 0.513 | 4217477 |
| 1 | WriterBenchmarks.OfficeIMO | RowCount=50000 | 15.193 ± 0.365 | 4234869 |

Ordinary OfficeIMO output includes coordinates and dimension; ordinary peer output omits them. Both are valid. Full finalization is timed and complete fields/date formatting/package structure are validated independently before timing. Different authored metadata explains a work boundary without making slower ordinary output invalid.

| Domain | Class / method | Native parameters | Mean ± SD (ms) | Allocated bytes/op |
| --- | --- | --- | ---: | ---: |
| 0 | CompactWriterBenchmarks.ExcelReaderWriter | RowCount=50000 | 9.610 ± 0.510 | 4216296 |
| 0 | CompactWriterBenchmarks.ExcelReaderWriterPrefetch | RowCount=50000 | 6.739 ± 0.314 | 4261965 |
| 0 | CompactWriterBenchmarks.OfficeIMOWithoutReferences | RowCount=50000 | 10.371 ± 0.268 | 4235560 |
| 1 | CompactWriterBenchmarks.ExcelReaderWriter | RowCount=50000 | 7.893 ± 0.163 | 4216413 |
| 1 | CompactWriterBenchmarks.ExcelReaderWriterPrefetch | RowCount=50000 | 4.441 ± 0.185 | 4220284 |
| 1 | CompactWriterBenchmarks.OfficeIMOWithoutReferences | RowCount=50000 | 7.309 ± 0.110 | 4235672 |

Coordinate-omission and shared-string policies stay explicit companions. Favorable compact results do not change the default-writer comparison or establish a win against the separately configured peer prefetch operation.

| Domain | Class / method | Native parameters | Mean ± SD (ms) | Allocated bytes/op |
| --- | --- | --- | ---: | ---: |
| 0 | NativeWriterBenchmarks.ExcelReaderWriter | Format=Xlsb&RowCount=50000 | 8.828 ± 0.304 | 4218488 |
| 0 | NativeWriterBenchmarks.OfficeIMOModelAndSave | Format=Xlsb&RowCount=50000 | 857.791 ± 48.627 | 461083288 |
| 1 | NativeWriterBenchmarks.ExcelReaderWriter | Format=Xlsb&RowCount=50000 | 5.743 ± 0.225 | 4216155 |
| 1 | NativeWriterBenchmarks.OfficeIMOModelAndSave | Format=Xlsb&RowCount=50000 | 619.232 ± 22.332 | 461084110 |

OfficeIMO native XLSB writing is supported through normal model construction/InsertObjects/Save. Its hundreds of milliseconds and roughly 461 MB cumulative allocation are substantial architecture costs versus the peer streamed writer. Typed/mapped callback/model fallbacks preserve the same supported-route costs; they are not automatic native streaming layouts. Existing spreadsheet writer-performance roadmap ownership is the follow-up location.

UTF-8 and Arrow writer companions use the original 100K × 4 prepared buffers/owned batch and reused 32 MiB output stream. Public row orchestration, batch/field access, validation/escaping and finalization are timed. There is no dedicated general RecordBatch writer. Styled 100K numeric rows use true named row defaults and resolved authored-cell XFs; complete stylesheet/reference/value proof stays part of Setup qualification.

Native BIFF8 XLS reading and model Save are supported. The peer manual/record XLS output is diagnostic: the retained required INDEX omission and installed Excel normal-open rejection prevent an equivalent-valid-output speed ranking. Sequential reader success alone cannot establish Excel conformance. No favorable ranking is obtained by treating unsupported/invalid output as a win.

### CSV sequential, typed and writer paths

| Domain | Class / method | Native parameters | Mean ± SD (ms) | Allocated bytes/op |
| --- | --- | --- | ---: | ---: |
| 0 | CsvTypedReadBenchmarks.ExcelReaderTyped | RowCount=50000 | 3.800 ± 0.798 | 4051779 |
| 0 | CsvTypedReadBenchmarks.OfficeIMORowsAs | RowCount=50000 | 8.014 ± 0.614 | 11449760 |
| 0 | CsvTypedReadBenchmarks.SylvanTyped | RowCount=50000 | 11.577 ± 0.857 | 11486693 |
| 1 | CsvTypedReadBenchmarks.ExcelReaderTyped | RowCount=50000 | 3.835 ± 0.203 | 4051702 |
| 1 | CsvTypedReadBenchmarks.OfficeIMORowsAs | RowCount=50000 | 5.820 ± 0.310 | 11451284 |
| 1 | CsvTypedReadBenchmarks.SylvanTyped | RowCount=50000 | 10.764 ± 1.099 | 11486688 |

Wide raw UTF-8/UTF-16 span work stays separate from materialized strings. Matching string-production cases can favor OfficeIMO, while typed and writer paths have valid losses. Composite RowCount/MaterializeStrings values and complete FullName are retained; do not merge False and True.

Public native CSV borrowing can return eligible normalized bytes already held by provider buffers. Decoded text/conversion fallback does not promise borrowing. General null/missing/schema behavior is not inferred from these populated numeric fixtures. Ordinary and record writers validate logical fields but use different formatting/line-ending output. Only the public UTF-8 writer's retained Setup proof establishes exact 6,900,000-byte output equality across engines; it does not hash every timed output.

### Original multi-million CSV and reductions

All methods consume the same full files:4.3M conversion-heavy 6-field rows (207,563,853B),8M narrow 3-integer rows (209,597,884B), and 3M conversion-heavy 6-field rows (144,811,299B). DOP 1 and 4 remain separate tables with full Input/Rows/Dop identity, field/order or direct consumed-field/multiplicity proof, and unchanged Huge job.

Core ordered automatic models, public decoded-text factory models, public text direct aggregate and public incremental asynchronous path/Stream aggregate perform different allocation/ownership work. Text-owned routes include whole-file read/decode inside timing. Async routes include full file I/O and pooled decoded-field snapshots; there is no hidden text preloading. Peer aggregate consumes UTF-8 ref/struct fields into native worker state. No fake equivalent native zero-copy OfficeIMO method fills that strategy difference.

The 8M async routes allocate approximately 2.26–2.27 GB cumulatively and take seconds; peer native reduction takes hundreds of milliseconds or less with around 1–2 MB cumulative managed allocation. Bounded queues/state do not imply low total allocation. DOP 4 can improve the peer native reduction while regressing OfficeIMO Core mapping; no universal parallel-scaling claim follows. The 12 per-group ReadAsync>0/Read=0 observations are untimed Setup probes of the OfficeIMO Stream aggregate only, including child instances timing peer/text-owned methods. They are not counters for timed invocations or evidence that those other methods use asynchronous I/O.

### Arrow output ownership and inference

| Domain | Class / method | Native parameters | Mean ± SD (ms) | Allocated bytes/op |
| --- | --- | --- | ---: | ---: |
| 0 | ArrowConversionBenchmarks.ExcelReader | RowCount=100000&Scenario=CsvAllString | 14.611 ± 1.127 | 17058621 |
| 0 | ArrowConversionBenchmarks.OfficeIMO | RowCount=100000&Scenario=CsvAllString | 42.747 ± 3.615 | 51803353 |
| 0 | ArrowConversionBenchmarks.ExcelReader | RowCount=100000&Scenario=CsvTyped | 9.691 ± 0.942 | 8528877 |
| 0 | ArrowConversionBenchmarks.OfficeIMO | RowCount=100000&Scenario=CsvTyped | 26.870 ± 1.481 | 22597304 |
| 0 | ArrowConversionBenchmarks.ExcelReader | RowCount=100000&Scenario=XlsbTyped | 15.154 ± 0.965 | 8532085 |
| 0 | ArrowConversionBenchmarks.OfficeIMO | RowCount=100000&Scenario=XlsbTyped | 24.170 ± 2.351 | 9529745 |
| 1 | ArrowConversionBenchmarks.ExcelReader | RowCount=100000&Scenario=CsvAllString | 11.420 ± 1.025 | 17058713 |
| 1 | ArrowConversionBenchmarks.OfficeIMO | RowCount=100000&Scenario=CsvAllString | 38.486 ± 2.103 | 51819047 |
| 1 | ArrowConversionBenchmarks.ExcelReader | RowCount=100000&Scenario=CsvTyped | 8.074 ± 0.828 | 8529309 |
| 1 | ArrowConversionBenchmarks.OfficeIMO | RowCount=100000&Scenario=CsvTyped | 22.283 ± 1.635 | 22598100 |
| 1 | ArrowConversionBenchmarks.ExcelReader | RowCount=100000&Scenario=XlsbTyped | 12.134 ± 0.379 | 8590061 |
| 1 | ArrowConversionBenchmarks.OfficeIMO | RowCount=100000&Scenario=XlsbTyped | 17.631 ± 1.004 | 9529518 |

Explicit typed CSV/XLSB uses required string/int64/microsecond timestamp(no timezone)/float64 fields; eight-string CSV uses required c0..c7. Each timed operation opens/parses the source, constructs one complete owned 100K-row RecordBatch, consumes its count and disposes readers/batches. Setup checks names/types/nullability/null counts/temporal units/every field/order. OfficeIMO is slower and allocates more in all three selected explicit cases in both domains.

The seventh peer-only inferred CSV row produces four required strings with parseText=false. OfficeIMO ordinary inference instead produces string/int32/microsecond timestamp/decimal128(29,10), separately validated. These outputs are not ranked as equivalent inference. Bounded 8,192 batches are qualified, separate unselected coverage; no timings are invented for them. Managed allocation is not total unmanaged Arrow memory.

### Encryption and causal Core attribution

Current encrypted scans pair complete integrity verification and every plaintext field/type/order. Original small ciphertext of 15,360 B with its authenticated 8,915 B plaintext oracle is preserved; the additional generated 300K × 4 input remains a distinct workload. MemoryInputFalse means FileStream for the small fixture and MemoryStream for the large prepared ciphertext. Whole-package eager buffering/authentication and decrypted package discovery remain included. Peer unverified/open-only/row-width originals remain diagnostics with no equivalent-work ratio.

The separate small-input intervention swaps only Core ba9→4ee with the other owners fixed at ba9, retaining a private harness source snapshot. It attributes the password-chain improvement, not the whole final candidate or all cipher/input sizes.

| Domain | Input route | Before mean ± SD (ms) | Core-only mean ± SD (ms) | Before bytes/op | Core-only bytes/op |
| --- | --- | ---: | ---: | ---: | ---: |
| 0 | FileStream | 189.878 ± 7.930 | 28.100 ± 2.276 | 144583008 | 572239 |
| 0 | byte[] | 182.793 ± 7.420 | 27.533 ± 1.549 | 144562916 | 552553 |
| 1 | FileStream | 176.548 ± 6.445 | 27.664 ± 1.202 | 144578908 | 570732 |
| 1 | byte[] | 170.381 ± 7.039 | 27.295 ± 1.405 | 144559036 | 550878 |

Those recorded means improve by 6.24–6.76× with approximately 99.6% less managed allocation for this small authenticated operation. The actual input, complete authentication/scanning, other-owner hashes, Core hashes, private harness and capture source are retained. They do not quantify the 300K effect or compare a mixed-owner bundle with the final peer candidate.

### Cold target use and diagnostic allocation

| Domain | Class / method | Native parameters | First-target-use mean ± SD (ms) | Second-use diagnostic allocated bytes/op |
| --- | --- | --- | ---: | ---: |
| 0 | ColdStartReadBenchmarks.ExcelReaderAttributes |  | 29.209 ± 4.338 | 21600 |
| 0 | ColdStartReadBenchmarks.OfficeIMOAutomaticMapping |  | 78.426 ± 49.660 | 96048 |
| 1 | ColdStartReadBenchmarks.ExcelReaderAttributes |  | 41.254 ± 8.753 | 136216 |
| 1 | ColdStartReadBenchmarks.OfficeIMOAutomaticMapping |  | 104.122 ± 59.560 | 115696 |

Cold latency targets the first mapping/layout operation in each fresh process after fixture/identity Setup. Setup has already exercised writers and ordinary metadata APIs; this is not first-ever API use. Full field/value proof is in a separate validation process to preserve target mapping/layout cold state.

MemoryDiagnoser obtains allocation through a subsequent extra-statistics invocation. The table therefore shows second-use diagnostic allocation, not first-use initialization allocation. This distinction is source-derived from pinned BDN 0.15.8 Engine.Run/GetExtraStats; no second-use latency is separately measured. Short/distribution warnings, individual launches and the significant domain 1 concurrent testhost CPU observation above are retained. Do not turn successful execution into an isolated-host cold ranking.

## Capabilities, omissions and reproduction

The comparison fills useful public OfficeIMO columns missing in retained upstream managed CSV, ADO, cold and other family packets. It does not imply every upstream internal method has a shipped counterpart. Existing public routes include ordinary automatic mapping, span-bearing RowsAsBorrowed factories, native XLS reading/model writing, native workbook/CSV UTF-8 borrowing, real asynchronous stream opening, decoded-record text/async direct reductions, declared streamed styles, UTF-8 row writing and optional Arrow export.

| Remaining boundary | Disposition |
| --- | --- |
| Arbitrary automatic ref-struct mapping |Public factory route exists; arbitrary automatic constructor/property mapping remains a product gap|
| Dedicated RecordBatch-to-XLSX writer |Public row/WriteUtf8 orchestration exists; dedicated general batch writer remains a product gap|
| Sep0.17.1 and SpreadCheetah1.28.0 columns |Missing comparison engines, not missing OfficeIMO APIs|
| Standalone SST index parser loops |Implementation microbenchmarks, not missing workbook-read APIs|
| Peer prefetch/fluent/inference policies |Specific strategies; ordinary supported public routes and unequal schema/work stay explicit|

Core/Excel/CSV public owner boundaries and target-framework details are recorded in the frozen capability inventory inside the archive. Source support does not establish that these APIs are present in published 3.4.4. Current product docs describe public contracts; existing Docs/ROADMAP.md owns open automatic borrowed mapping, Arrow writer and relevant performance follow-ups, rather than another backlog.

The isolated suite targets .NET 10 and remains outside the normal solution. Comparison dependencies are test-only: ExcelReader.NET/Arrow 6.0.0 (MIT, pinned ca5b50f), Sylvan.Data.Excel 0.5.8, Sylvan.Data 0.2.17 and Sylvan.Data.Csv 1.4.4 (MIT); BenchmarkDotNet 0.15.8 and Apache.Arrow 23.0.0 retain their licenses/notices. Fixture copyright/producer attribution remains separate from package licenses. Real/encrypted fixtures are hash-referenced rather than redistributed in this packet. The small upstream encryption producer is unpinned; independent msoffcrypto authentication/decryption and installed Excel evidence qualify its oracle.

Use [the suite README](../../Benchmarks/ExcelReaderTyped/README.md) and [third-party notices](../../Benchmarks/ExcelReaderTyped/THIRD-PARTY-NOTICES.md) for source/package/saved selectors, optional flags, fixture generation and full validation commands. Before timing, run --validate-bdn plus relevant full-field commands and an actual generated-process Dry qualification. Reproduce a selected contract with the pinned source/runtime inputs, matching output policies, actual fixture hashes, complete integrity settings, affinity/priority and job recipe. Archived contexts retain exact original argument arrays and source/runtime metadata. Machine-specific scratch paths in native logs are recorded provenance, not portable defaults.

Frozen qualification passed 6,149 Excel cases with five existing opt-in skips and zero failures, and 801 CSV cases. Windows focus covered 86 net10/86 net8/69 net472; all four owner TFMs compiled without warnings/errors. The suite declaration validator covers 59 types/170 methods/477 parameter cases; the selected measured matrix is 306 per domain. A generated-reference compiler failure initially produced zero measurements; the exact saved-reference hook repair subsequently qualified 136 eligibility cases and the measured resume. Failure/Dry records remain unranked and are not pooled into 730.

The archive preserves canonical native BDN/PowerForge outputs and immutable contexts/imports/sidecars, exact source/runtime/harness/fixture identities, failure/eligibility proof, and historical classifications. Its inventory/extraction hash proof explains exactly what is retained. Later upstream integration, local PR partition or package builds receive separate identities; they do not relabel the measured 830951 evidence. This report makes no package-publication, external deployment or universal performance claim.