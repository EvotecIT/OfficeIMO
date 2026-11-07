#if NET10_0_OR_GREATER
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Complete packages from equivalent generated records, with serial and parallel work labelled explicitly.</summary>
[MemoryDiagnoser]
public class ExcelWriterOptionsComparisonBenchmarks {
    [Params(ExcelFileFormat.Xlsx, ExcelFileFormat.Xlsb)]
    public ExcelFileFormat Format { get; set; }

    [Params(50_000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) BenchmarkProcessorAffinity.ApplyPriority(priority);
        var validator = new ExcelReaderGeneratedWriteBenchmarks { Format = Format, RowCount = RowCount };
        validator.Validate(OfficeIMO_Serial());
        validator.Validate(ExcelReaderNet_Serial());
        validator.Validate(ExcelReaderNet_Prefetch());
        validator.Validate(ExcelReaderNet_Parallel1());
        validator.Validate(ExcelReaderNet_Parallel8());
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO_Serial() => ExcelReaderGeneratedBenchmarks.WriteOfficeIMO(Format, RowCount);

    [Benchmark]
    public byte[] ExcelReaderNet_Serial() => ExcelReaderGeneratedBenchmarks.WriteExcelReader(Format, RowCount);

    [Benchmark]
    public byte[] ExcelReaderNet_Prefetch() => ExcelReaderGeneratedBenchmarks.WriteExcelReader(Format, RowCount, prefetch: true);

    [Benchmark]
    public byte[] ExcelReaderNet_Parallel1() => ExcelReaderGeneratedBenchmarks.WriteExcelReader(Format, RowCount, parallelism: 1);

    [Benchmark]
    public byte[] ExcelReaderNet_Parallel8() => ExcelReaderGeneratedBenchmarks.WriteExcelReader(Format, RowCount, parallelism: 8);
}
#endif
