using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Tabular.Benchmarks;

[MemoryDiagnoser]
[BenchmarkCategory("CSV", "String", "Pinned65K")]
public class CsvStringBenchmarks {
    [GlobalSetup]
    public void Setup() => FixtureData.EnsureAuthentic();

    [Benchmark(Baseline = true)]
    public Observation Sylvan() => TabularBenchmarkOperations.ReadSylvanCsvStrings();

    [Benchmark]
    public Observation OfficeIMO() => TabularBenchmarkOperations.ReadOfficeCsvStrings();
}

[MemoryDiagnoser]
[BenchmarkCategory("CSV", "TypedManual", "Pinned65K")]
public class CsvTypedBenchmarks {
    [GlobalSetup]
    public void Setup() => FixtureData.EnsureAuthentic();

    [Benchmark(Baseline = true)]
    public Observation Sylvan() => TabularBenchmarkOperations.ReadSylvanCsvTyped();

    [Benchmark]
    public Observation OfficeIMO() => TabularBenchmarkOperations.ReadOfficeCsvTyped();
}

[MemoryDiagnoser]
[BenchmarkCategory("XLSX", "TypedManual", "Pinned65K")]
public class XlsxTypedBenchmarks {
    [GlobalSetup]
    public void Setup() => FixtureData.EnsureAuthentic();

    [Benchmark(Baseline = true)]
    public Observation Sylvan() => TabularBenchmarkOperations.ReadSylvanXlsxTyped();

    [Benchmark]
    public Observation OfficeIMO() => TabularBenchmarkOperations.ReadOfficeXlsxTyped();
}

[MemoryDiagnoser]
[BenchmarkCategory("XLSX", "TypedBinder", "Pinned65K")]
public class XlsxBinderBenchmarks {
    [GlobalSetup]
    public void Setup() => FixtureData.EnsureAuthentic();

    [Benchmark(Baseline = true)]
    public Observation Sylvan() => TabularBenchmarkOperations.ReadSylvanXlsxRecords();

    [Benchmark]
    public Observation OfficeIMO() => TabularBenchmarkOperations.ReadOfficeXlsxRecords();
}

[MemoryDiagnoser]
[BenchmarkCategory("XLSB", "TypedManual", "Pinned65K")]
public class XlsbTypedBenchmarks {
    [GlobalSetup]
    public void Setup() => FixtureData.EnsureAuthentic();

    [Benchmark(Baseline = true)]
    public Observation Sylvan() => TabularBenchmarkOperations.ReadSylvanXlsbTyped();

    [Benchmark]
    public Observation OfficeIMO() => TabularBenchmarkOperations.ReadOfficeXlsbTyped();
}
