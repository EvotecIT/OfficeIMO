#if NET10_0_OR_GREATER
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using OfficeIMO.Data;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;
using GeneratedRecord = OfficeIMO.Excel.Benchmarks.ExcelReaderGeneratedBenchmarks.GeneratedRecord;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Compares complete, ordered typed reads using each library's requested parallelism.</summary>
[MemoryDiagnoser]
public class ExcelOrderedParallelComparisonBenchmarks {
    private static readonly ExcelParser<GeneratedRecord> Parser = ExcelParser.FromAttributes<GeneratedRecord>();
    private ExcelReaderGeneratedBenchmarks? _fixture;
    private ParallelRowMappingOptions _mapping = null!;
    private SheetParallelOptions _parallel = null!;
    private string _path = string.Empty;
    private long _expected;
    private bool _validate;

    [Params(ExcelFileFormat.Xlsx, ExcelFileFormat.Xlsb)]
    public ExcelFileFormat Format { get; set; }

    [Params(1, 4)]
    public int RequestedDop { get; set; }

    [Params(false, true)]
    public bool EscapedNames { get; set; }

    [Params(false, true)]
    public bool InferSchema { get; set; }

    [GlobalSetup]
    public void Setup() {
        _fixture = new ExcelReaderGeneratedBenchmarks {
            Format = Format, RowCount = 50_000, EscapedNames = EscapedNames
        };
        _fixture.Setup();
        _path = _fixture.InputPath;
        _mapping = new ParallelRowMappingOptions { MaxDegreeOfParallelism = RequestedDop, BatchSize = 1024 };
        _parallel = new SheetParallelOptions { DegreeOfParallelism = RequestedDop };
        for (int index = 0; index < 50_000; index++)
            _expected = ExcelReaderGeneratedBenchmarks.Add(_expected,
                ExcelReaderGeneratedBenchmarks.CreateRecord(index, EscapedNames));
        _validate = true;
        try {
            OfficeIMO_Ordered();
            ExcelReaderNet_Ordered().GetAwaiter().GetResult();
        } catch { Cleanup(); throw; }
        finally { _validate = false; }
    }

    [GlobalCleanup]
    public void Cleanup() => _fixture?.Cleanup();

    [Benchmark(Baseline = true)]
    public long OfficeIMO_Ordered() {
        using var reader = ExcelDocument.OpenReader(_path, new ExcelReaderOptions { InferSchema = InferSchema });
        long checksum = 0;
        int count = 0;
        foreach (GeneratedRecord row in reader.RowsAsParallel<GeneratedRecord>(_mapping)) {
            Validate(count++, row);
            checksum = ExcelReaderGeneratedBenchmarks.Add(checksum, row);
        }
        return Check(checksum, count);
    }

    [Benchmark]
    public async Task<long> ExcelReaderNet_Ordered() {
        using IExcelWorkbook workbook = Format switch {
            ExcelFileFormat.Xlsx => ExcelReaderApi.FromXlsxFile(_path),
            ExcelFileFormat.Xlsb => ExcelReaderApi.FromXlsbFile(_path),
            _ => throw new ArgumentOutOfRangeException()
        };
        long checksum = 0;
        int count = 0;
        await foreach (GeneratedRecord row in SheetParallel.ParseAsync(workbook.FirstSheet, Parser, _parallel)) {
            Validate(count++, row);
            checksum = ExcelReaderGeneratedBenchmarks.Add(checksum, row);
        }
        return Check(checksum, count);
    }

    private void Validate(int index, GeneratedRecord row) {
        if (_validate) ExcelReaderGeneratedBenchmarks.Validate(
            index, row.Name, row.Id, row.CreatedOn, row.Amount, EscapedNames);
    }

    private long Check(long checksum, int count) {
        if (count != 50_000 || checksum != _expected)
            throw new InvalidDataException($"Incomplete ordered {Format} read: {count} rows, {checksum}.");
        return checksum;
    }
}
#endif
