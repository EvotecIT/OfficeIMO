using System.Globalization;
using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>Reads the upstream SST-heavy input with eight text columns and three typed value columns.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("StringHeavyOriginal")]
public class StringHeavyReadBenchmarks {
    private readonly WorkbookScanWorkload _workload = new();

    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 65_536;

    [Params(ComparisonWorkbookFormat.Xlsx, ComparisonWorkbookFormat.Xlsb)]
    public ComparisonWorkbookFormat Format { get; set; }

    public IEnumerable<int> RowCounts() {
        string value = Environment.GetEnvironmentVariable("OFFICEIMO_STRING_BENCHMARK_ROWS") ?? "65536";
        foreach (string item in value.Split(',', StringSplitOptions.TrimEntries)) {
            if (!int.TryParse(item, NumberStyles.None, CultureInfo.InvariantCulture, out int count) || count is < 1 or > 1_000_000)
                throw new ArgumentException("OFFICEIMO_STRING_BENCHMARK_ROWS requires counts between 1 and 1000000.");
            yield return count;
        }
    }

    [GlobalSetup]
    public async Task SetupAsync() {
        BenchmarkInput.WriteDescription();
        byte[] bytes = Format switch {
            ComparisonWorkbookFormat.Xlsx => await StringHeavyWorkbookGenerator.BuildXlsxAsync(RowCount),
            ComparisonWorkbookFormat.Xlsb => await StringHeavyWorkbookGenerator.BuildXlsbAsync(RowCount),
            _ => throw new ArgumentOutOfRangeException(nameof(Format)),
        };
        object[] headers = StringHeavyWorkbookGenerator.Headers.Cast<object>().ToArray();
        _workload.Setup(bytes, Format, RowCount + 1, headers.Length,
            index => index == 0 ? headers : StringHeavyWorkbookGenerator.ExpectedValues(index));
    }

    [Benchmark(Baseline = true)]
    public long ExcelReaderOriginal() => _workload.ExcelReader();
    [Benchmark]
    public long ExcelReaderPrefetch() => _workload.ExcelReader(prefetch: true);
    [Benchmark]
    public long ExcelReaderMemory() => _workload.ExcelReader(memory: true);
    [Benchmark]
    public long OfficeIMOBytes() => _workload.OfficeIMO();
    [Benchmark]
    public long OfficeIMOStream() => _workload.OfficeIMO(streamInput: true);
    [Benchmark]
    public long Sylvan() => _workload.Sylvan();

    internal long ExcelReaderStringsMaterialized(bool intern = false) => _workload.ExcelReader(materializeStrings: true, intern: intern);
}

/// <summary>Consumes materialized string cells while preserving the peer's explicit cache-policy case.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("StringHeavyStringsMaterialized")]
public class MaterializedStringHeavyReadBenchmarks {
    private readonly StringHeavyReadBenchmarks _workload = new();

    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 65_536;

    [Params(ComparisonWorkbookFormat.Xlsx, ComparisonWorkbookFormat.Xlsb)]
    public ComparisonWorkbookFormat Format { get; set; }

    public IEnumerable<int> RowCounts() => _workload.RowCounts();

    [GlobalSetup]
    public Task SetupAsync() { _workload.RowCount = RowCount; _workload.Format = Format; return _workload.SetupAsync(); }

    [Benchmark(Baseline = true)]
    public long ExcelReaderStringsMaterialized() => _workload.ExcelReaderStringsMaterialized();
    [Benchmark]
    public long ExcelReaderStringsInterned() => _workload.ExcelReaderStringsMaterialized(intern: true);
    [Benchmark]
    public long OfficeIMOBytes() => _workload.OfficeIMOBytes();
    [Benchmark]
    public long OfficeIMOStream() => _workload.OfficeIMOStream();
    [Benchmark]
    public long Sylvan() => _workload.Sylvan();
}
