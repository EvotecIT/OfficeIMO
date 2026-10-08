using System.Text;
using System.Runtime.InteropServices;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader.Csv;
using OfficeIMO.CSV;
using OfficeIMO.Data;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

public ref struct CsvHeavyRefRecord {
    public ReadOnlySpan<byte> Region { get; set; }
    public ReadOnlySpan<byte> Country { get; set; }
    public DateTime OrderDate { get; set; }
    public decimal UnitPrice { get; set; }
    public decimal TotalRevenue { get; set; }
    public int Units { get; set; }
}

[StructLayout(LayoutKind.Auto)]
public struct CsvNarrowStructRecord {
    public int A { get; set; }
    public int B { get; set; }
    public int C { get; set; }
}

internal readonly record struct CsvAggregateValue(long Value);

/// <summary>Retains caller reduction over mapped results as a separate diagnostic.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("CsvCallerReductionDiagnostic")]
public class CsvParallelAggregateBenchmarks {
    private CsvParallelFixture _fixture = null!;
#if OFFICEIMO_BENCHMARK_NEW_APIS
    private bool _validate;
#endif
    // BDN runs operations serially within one process. Accumulator construction
    // captures this setup state before worker processing starts.
    private static CsvParallelFixture? _validationFixture;
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
        Console.WriteLine($"CSV aggregate {Input}: bytes={_fixture.Bytes}, sha256={_fixture.Fingerprint}, DOP={Dop}, expectedSum={_fixture.ExpectedSum}; all fields parsed.");
#if OFFICEIMO_BENCHMARK_NEW_APIS
        _validate = true;
#endif
        _validationFixture = _fixture;
        try {
            await ExcelReaderAggregate();
#if OFFICEIMO_BENCHMARK_NEW_APIS
            OfficeIMOCallerReduction();
#endif
        }
        catch { Cleanup(); throw; }
        finally {
#if OFFICEIMO_BENCHMARK_NEW_APIS
            _validate = false;
#endif
            _validationFixture = null;
        }
    }
    [GlobalCleanup]
    public void Cleanup() => _fixture?.Dispose();

    [Benchmark]
    public async Task<long> ExcelReaderAggregate() {
        var options = new CsvParallelOptions { DegreeOfParallelism = Dop, HeaderRow = 1 };
        if (_fixture.Heavy) {
            HeavyAccumulator result = await CsvParallel.AggregateAsync<HeavyAccumulator, CsvHeavyRefRecord>(
                _fixture.Path, ExcelParser.FromAttributes<CsvHeavyRefRecord>(), options);
            return _fixture.Check(result.Sum, result.Count);
        }
        NarrowAccumulator narrow = await CsvParallel.AggregateAsync<NarrowAccumulator, CsvNarrowStructRecord>(
            _fixture.Path, ExcelParser.FromAttributes<CsvNarrowStructRecord>(), options);
        return _fixture.Check(narrow.Sum, narrow.Count);
    }

#if OFFICEIMO_BENCHMARK_NEW_APIS
    [Benchmark]
    public long OfficeIMOCallerReduction() {
        string text = File.ReadAllText(_fixture.Path);
        var loadOptions = new CsvLoadOptions { MaxInputBytes = Math.Max(CsvLoadOptions.DefaultMaxInputBytes, _fixture.Bytes) };
        var parallelOptions = new ParallelRowMappingOptions { MaxDegreeOfParallelism = Dop == 0 ? null : Dop };
        IEnumerable<CsvAggregateValue> values = _fixture.Heavy
            ? CsvDocument.ReadTextRowsAsParallel(text, HeavyFactory, loadOptions: loadOptions, parallelOptions: parallelOptions)
            : CsvDocument.ReadTextRowsAsParallel(text, NarrowFactory, loadOptions: loadOptions, parallelOptions: parallelOptions);
        long sum = 0;
        int count = 0;
        foreach (CsvAggregateValue value in values) { sum += value.Value; count++; }
        return _fixture.Check(sum, count);
    }

    private CsvRecordFactory<CsvAggregateValue> HeavyFactory(CsvRecordHeader header) {
        int region = header.GetOrdinal("Region"), country = header.GetOrdinal("Country"), date = header.GetOrdinal("OrderDate"),
            price = header.GetOrdinal("UnitPrice"), revenue = header.GetOrdinal("TotalRevenue"), units = header.GetOrdinal("Units");
        return row => {
            ReadOnlySpan<char> regionText = row.GetSpan(region), countryText = row.GetSpan(country);
            DateTime orderDate = row.GetDateTime(date);
            decimal unitPrice = row.GetDecimal(price), totalRevenue = row.GetDecimal(revenue);
            int count = row.GetInt32(units);
            if (_validate) _fixture.ValidateHeavyParts(regionText, countryText, orderDate, unitPrice, totalRevenue, count);
            return new CsvAggregateValue(count);
        };
    }

    private CsvRecordFactory<CsvAggregateValue> NarrowFactory(CsvRecordHeader header) {
        int a = header.GetOrdinal("A"), b = header.GetOrdinal("B"), c = header.GetOrdinal("C");
        return row => {
            int first = row.GetInt32(a), second = row.GetInt32(b), third = row.GetInt32(c);
            if (_validate && (first < 0 || first >= Input.Rows || second != first * 3 || third != first * 7))
                throw new InvalidDataException("Aggregate integer fields differ.");
            return new CsvAggregateValue(first);
        };
    }
#endif

    private sealed class HeavyAccumulator : ICsvAccumulator<HeavyAccumulator, CsvHeavyRefRecord> {
        private readonly CsvParallelFixture? _validation = _validationFixture;
        public long Sum { get; private set; }
        public int Count { get; private set; }
        public void Add(CsvHeavyRefRecord row) {
            if (_validation is not null) _validation.ValidateHeavyParts(Encoding.UTF8.GetString(row.Region),
                Encoding.UTF8.GetString(row.Country), row.OrderDate, row.UnitPrice, row.TotalRevenue, row.Units);
            Sum += row.Units; Count++;
        }
        public void Merge(HeavyAccumulator following) { Sum += following.Sum; Count += following.Count; }
    }

    private sealed class NarrowAccumulator : ICsvAccumulator<NarrowAccumulator, CsvNarrowStructRecord> {
        private readonly CsvParallelFixture? _validation = _validationFixture;
        public long Sum { get; private set; }
        public int Count { get; private set; }
        public void Add(CsvNarrowStructRecord row) {
            if (_validation is not null) {
                if (row.A < 0 || row.A >= _validation.Input.Rows) throw new InvalidDataException("Aggregate integer row index differs.");
                _validation.Validate(new CsvNarrowRecord { A = row.A, B = row.B, C = row.C }, row.A);
            }
            Sum += row.A; Count++;
        }
        public void Merge(NarrowAccumulator following) { Sum += following.Sum; Count += following.Count; }
    }
}
