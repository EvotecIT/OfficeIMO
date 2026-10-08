using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>Reads the exact hash-pinned workbook bytes in the upstream real-data suite.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("RealDataOriginal")]
public class RealDataReadBenchmarks {
    private readonly WorkbookScanWorkload _workload = new();

    [ParamsAllValues]
    public ComparisonWorkbookFormat Format { get; set; }

    [GlobalSetup]
    public void Setup() {
        BenchmarkInput.WriteDescription();
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        string name = "65K_Records_Data." + Format.ToString().ToLowerInvariant();
        string expectedHash;
        if (Format == ComparisonWorkbookFormat.Xlsm) {
            expectedHash = "6E4B3A60D3C59C075BC53E4AF6DAE482DEC8532A0AF0309DA4EFB6AD59F39ABE";
            MarkPflug65KFixture.EnsureAuthentic(name, expectedHash, new Uri(
                "https://raw.githubusercontent.com/GabrielMarquezMatte/ExcelReader/ca5b50f99e8ef57ab476f0a2bc8043558d58b28d/tests/ExcelReader.Benchmarks/Data/" + name));
        } else {
            MarkPflug65KFixture.EnsureAuthentic(name);
            expectedHash = MarkPflug65KFixture.GetHashes()[name];
        }
        byte[] bytes = File.ReadAllBytes(Path.Combine(MarkPflug65KFixture.Root, name));
        Console.WriteLine($"Real-data fixture: {name}, SHA256={expectedHash}; XLS/XLSX/XLSB match the pinned upstream bytes.");
        _workload.Setup(bytes, Format, MarkPflug65KFixture.ExpectedRows + 1, MarkPflug65KFixture.ExpectedColumns);
    }

    [Benchmark(Baseline = true)]
    public long ExcelReaderOriginal() => _workload.ExcelReader();
    [Benchmark]
    public long ExcelReaderMemory() => _workload.ExcelReader(memory: true);
    [Benchmark]
    public long OfficeIMOBytes() => _workload.OfficeIMO();
    [Benchmark]
    public long OfficeIMOStream() => _workload.OfficeIMO(streamInput: true);
    [Benchmark]
    public long Sylvan() => _workload.Sylvan();

    internal long ExcelReaderStringsMaterialized() => _workload.ExcelReader(materializeStrings: true);
    internal long ExcelReaderPrefetch(bool memory) => _workload.ExcelReader(prefetch: true, memory: memory);
}

/// <summary>Preserves the upstream ZIP-format decompression-prefetch cases separately from defaults.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("RealDataPrefetch")]
public class PrefetchedRealDataReadBenchmarks {
    private readonly RealDataReadBenchmarks _workload = new();

    [Params(ComparisonWorkbookFormat.Xlsx, ComparisonWorkbookFormat.Xlsm, ComparisonWorkbookFormat.Xlsb)]
    public ComparisonWorkbookFormat Format { get; set; }

    [GlobalSetup]
    public void Setup() { _workload.Format = Format; _workload.Setup(); }

    [Benchmark(Baseline = true)]
    public long ExcelReaderPrefetch() => _workload.ExcelReaderPrefetch(memory: false);
    [Benchmark]
    public long ExcelReaderMemoryPrefetch() => _workload.ExcelReaderPrefetch(memory: true);
    [Benchmark]
    public long OfficeIMOBytes() => _workload.OfficeIMOBytes();
    [Benchmark]
    public long Sylvan() => _workload.Sylvan();
}

/// <summary>Separately consumes materialized string cells in the real-data fixture.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("RealDataStringsMaterialized")]
public class MaterializedRealDataReadBenchmarks {
    private readonly RealDataReadBenchmarks _workload = new();

    [ParamsAllValues]
    public ComparisonWorkbookFormat Format { get; set; }

    [GlobalSetup]
    public void Setup() { _workload.Format = Format; _workload.Setup(); }

    [Benchmark(Baseline = true)]
    public long ExcelReaderStringsMaterialized() => _workload.ExcelReaderStringsMaterialized();
    [Benchmark]
    public long OfficeIMOBytes() => _workload.OfficeIMOBytes();
    [Benchmark]
    public long OfficeIMOStream() => _workload.OfficeIMOStream();
    [Benchmark]
    public long Sylvan() => _workload.Sylvan();
}
