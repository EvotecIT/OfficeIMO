#if NET10_0_OR_GREATER
using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Compares complete materialized scans with optional ZIP decompression prefetch.</summary>
[MemoryDiagnoser]
public class ExcelReaderGeneratedZipPrefetchBenchmarks {
    private ExcelReaderGeneratedBenchmarks? _benchmark;

    [Params(ExcelFileFormat.Xlsx, ExcelFileFormat.Xlsb)]
    public ExcelFileFormat Format { get; set; }

    [Params(50_000)]
    public int RowCount { get; set; }

    [Params(false, true)]
    public bool EscapedNames { get; set; }

    [GlobalSetup]
    public void Setup() {
        _benchmark = new ExcelReaderGeneratedBenchmarks {
            Format = Format, RowCount = RowCount, EscapedNames = EscapedNames
        };
        _benchmark.Setup();
    }

    [GlobalCleanup]
    public void Cleanup() => _benchmark?.Cleanup();

    [Benchmark(Baseline = true)]
    public long OfficeIMO() => _benchmark!.OfficeIMO_Read();

    [Benchmark]
    public long ExcelReaderNet() => _benchmark!.ReadExcelReader(options: null);

    [Benchmark]
    public long ExcelReaderNet_Prefetch() =>
        _benchmark!.ReadExcelReader(ExcelReaderGeneratedBenchmarks.PrefetchReaderOptions);
}
#endif
