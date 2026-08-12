# Streaming `IDataReader` table write evidence

This 2026-08-10 run measures a 25,000-row, eight-column `IDataReader` export to
a styled XLSX table. Every implementation must produce the same typed cell
sequence, pass `OpenXmlValidator`, and include a worksheet table with an
AutoFilter. Styled ranges without a table definition do not qualify for this
lane.

The run used .NET 10.0.10 in Release mode on Windows 10.0.26200.0 and an AMD
Ryzen 9 9950X3D2 16-Core Processor. Each process used High priority, twelve
warmups, 31 measured iterations, and one fixed logical processor. Allocation is
the average managed allocation per operation.

| Affinity | Library | Mean | Median | Allocation | Package size |
| --- | --- | ---: | ---: | ---: | ---: |
| `0x1` | OfficeIMO.Excel | 24.22 ms | 22.78 ms | 6,347.7 KB | 908,255 B |
| `0x1` | ClosedXML 0.105.0 | 478.95 ms | 490.34 ms | 170,715.7 KB | 1,155,504 B |
| `0x1` | EPPlus 8.6.3 | 427.70 ms | 421.41 ms | 117,311.1 KB | 1,117,051 B |
| `0x10000` | OfficeIMO.Excel | 18.24 ms | 16.26 ms | 6,347.7 KB | 908,255 B |
| `0x10000` | ClosedXML 0.105.0 | 355.44 ms | 346.78 ms | 170,715.7 KB | 1,155,504 B |
| `0x10000` | EPPlus 8.6.3 | 356.93 ms | 345.28 ms | 117,311.2 KB | 1,117,051 B |

OfficeIMO had the lowest observed mean, median, allocation, and package size on
both processor domains. Its mean was 17.7-19.8 times lower than the other valid
table implementations, while allocation was 18.5-26.9 times lower.

An immediately preceding run of the former OfficeIMO worksheet-materialization
path used the same machine, data, affinity, priority, warmups, and iteration
count. It measured 55.80 ms and 13,480.8 KB on `0x1`, and 48.05 ms and
13,582.3 KB on `0x10000`. The streaming table path reduced observed mean time by
56.6% and 62.0%, and average allocation by 53.0% and 53.3%, respectively. These
are workstation observations, not universal constants.

Reproduce the current comparison with:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare .\datareader-table-cpu0.json --rows 25000 --scenario write-datareader-table --skip-legacy-epplus --warmup 12 --iterations 31 --affinity 0x1 --priority High --library OfficeIMO.Excel,ClosedXML,EPPlus
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Excel.Benchmarks\OfficeIMO.Excel.Benchmarks.csproj -- compare .\datareader-table-cpu16.json --rows 25000 --scenario write-datareader-table --skip-legacy-epplus --warmup 12 --iterations 31 --affinity 0x10000 --priority High --library OfficeIMO.Excel,ClosedXML,EPPlus
```

The runner validates output before timing. MiniExcel 1.45.0 is not included in
this lane because its table-style configuration did not emit an XLSX table
definition in the preflight workbook; it remains eligible for plain worksheet
write comparisons.
