using System.Globalization;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader.Csv;
using OfficeIMO.CSV;
using OfficeIMO.Data;
using Sylvan.Data;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>Full-file typed parallel reads, including source decoding for decoded-text routes.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("CsvParallelTyped")]
public class CsvParallelTypedBenchmarks {
    private CsvParallelFixture _fixture = null!;
    private bool _validate;
    [ParamsSource(nameof(Cases))]
    public CsvParallelInput Input { get; set; }
    [ParamsSource(nameof(Degrees))]
    public int Dop { get; set; }
    public IEnumerable<CsvParallelInput> Cases() => CsvParallelFixture.Cases();
    public IEnumerable<int> Degrees() => CsvParallelFixture.Degrees();

    [GlobalSetup]
    public async Task SetupAsync() {
        _fixture = new(Input);
        CsvBenchmarkFixture.DescribeEngines();
        Console.WriteLine($"CSV parallel {Input}: bytes={_fixture.Bytes}, sha256={_fixture.Fingerprint}, DOP={Dop}, expectedSum={_fixture.ExpectedSum}; full-file input/decoding is timed.");
        _validate = true;
        try {
            await ExcelReaderParallel(); ExcelReaderSequential(); OfficeIMOCoreParallel(); SylvanSequential();
#if OFFICEIMO_BENCHMARK_NEW_APIS
            OfficeIMOTextParallel();
#endif
        } catch { Cleanup(); throw; }
        finally { _validate = false; }
    }
    [GlobalCleanup]
    public void Cleanup() => _fixture?.Dispose();

    [Benchmark(Baseline = true)]
    public async Task<long> ExcelReaderParallel() {
        int count = 0;
        long sum = 0;
        var options = new CsvParallelOptions { DegreeOfParallelism = Dop };
        if (_fixture.Heavy) {
            await foreach (CsvHeavyRecord row in CsvParallel.ParseAsync(_fixture.Path, ExcelParser.FromAttributes<CsvHeavyRecord>(), options)) {
                if (_validate) _fixture.Validate(row, count);
                sum += row.Units; count++;
            }
        } else {
            await foreach (CsvNarrowRecord row in CsvParallel.ParseAsync(_fixture.Path, ExcelParser.FromAttributes<CsvNarrowRecord>(), options)) {
                if (_validate) _fixture.Validate(row, count);
                sum += row.A; count++;
            }
        }
        return _fixture.Check(sum, count);
    }

    [Benchmark]
    public long ExcelReaderSequential() {
        using var workbook = ExcelReaderApi.FromCsvFile(_fixture.Path);
        return _fixture.Heavy
            ? Consume(ExcelParser.FromAttributes<CsvHeavyRecord>().Parse(workbook.FirstSheet))
            : Consume(ExcelParser.FromAttributes<CsvNarrowRecord>().Parse(workbook.FirstSheet));
    }

    [Benchmark]
    public long OfficeIMOCoreParallel() {
        using var reader = CsvDocument.OpenDataReader(_fixture.Path, LoadOptions());
        var options = ParallelOptions();
        return _fixture.Heavy ? Consume(reader.RowsAsParallel<CsvHeavyRecord>(options))
            : Consume(reader.RowsAsParallel<CsvNarrowRecord>(options));
    }

#if OFFICEIMO_BENCHMARK_NEW_APIS
    [Benchmark]
    public long OfficeIMOTextParallel() {
        // The public span API accepts decoded text. File I/O and UTF-8 decoding belong
        // to this operation, rather than giving one library a pre-decoded setup input.
        string text = File.ReadAllText(_fixture.Path);
        return _fixture.Heavy ? Consume(CsvDocument.ReadTextRowsAsParallel(text, HeavyFactory,
            loadOptions: LoadOptions(), parallelOptions: ParallelOptions()))
            : Consume(CsvDocument.ReadTextRowsAsParallel(text, NarrowFactory,
                loadOptions: LoadOptions(), parallelOptions: ParallelOptions()));
    }
#endif

    [Benchmark]
    public long SylvanSequential() {
        using var text = new StreamReader(_fixture.Path);
        using var reader = global::Sylvan.Data.Csv.CsvDataReader.Create(text,
            new global::Sylvan.Data.Csv.CsvDataReaderOptions { Culture = CultureInfo.InvariantCulture });
        return _fixture.Heavy ? Consume(reader.GetRecords<CsvHeavyRecord>()) : Consume(reader.GetRecords<CsvNarrowRecord>());
    }

    private long Consume(IEnumerable<CsvHeavyRecord> records) {
        long sum = 0;
        int count = 0;
        foreach (CsvHeavyRecord row in records) {
            if (_validate) _fixture.Validate(row, count);
            sum += row.Units; count++;
        }
        return _fixture.Check(sum, count);
    }
    private long Consume(IEnumerable<CsvNarrowRecord> records) {
        long sum = 0;
        int count = 0;
        foreach (CsvNarrowRecord row in records) {
            if (_validate) _fixture.Validate(row, count);
            sum += row.A; count++;
        }
        return _fixture.Check(sum, count);
    }

    private CsvLoadOptions LoadOptions() => new() { MaxInputBytes = Math.Max(CsvLoadOptions.DefaultMaxInputBytes, _fixture.Bytes) };
    private ParallelRowMappingOptions ParallelOptions() => new() { MaxDegreeOfParallelism = Dop == 0 ? null : Dop };

#if OFFICEIMO_BENCHMARK_NEW_APIS
    private static CsvRecordFactory<CsvHeavyRecord> HeavyFactory(CsvRecordHeader header) {
        int region = header.GetOrdinal("Region"), country = header.GetOrdinal("Country"), date = header.GetOrdinal("OrderDate"),
            price = header.GetOrdinal("UnitPrice"), revenue = header.GetOrdinal("TotalRevenue"), units = header.GetOrdinal("Units");
        return row => new CsvHeavyRecord { Region = row.GetString(region), Country = row.GetString(country), OrderDate = row.GetDateTime(date),
            UnitPrice = row.GetDecimal(price), TotalRevenue = row.GetDecimal(revenue), Units = row.GetInt32(units) };
    }
    private static CsvRecordFactory<CsvNarrowRecord> NarrowFactory(CsvRecordHeader header) {
        int a = header.GetOrdinal("A"), b = header.GetOrdinal("B"), c = header.GetOrdinal("C");
        return row => new CsvNarrowRecord { A = row.GetInt32(a), B = row.GetInt32(b), C = row.GetInt32(c) };
    }
#endif
}
