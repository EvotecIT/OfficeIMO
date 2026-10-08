#if OFFICEIMO_BENCHMARK_NEW_APIS && OFFICEIMO_BENCHMARK_CSV
using System.Text;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader.Csv;
using OfficeIMO.CSV;
using OfficeIMO.Data;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>Reduces the original CSV fields into independent states with complete input work timed.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("CsvDirectAggregate")]
public class CsvDirectAggregateBenchmarks {
    private CsvParallelFixture _fixture = null!;
    private CsvAggregateProof? _proof;
    // The original peer API constructs accumulator states through parameterless
    // constructors. Setup activates proof only for one serial benchmark operation.
    private static CsvAggregateProof? _peerProof;
    [ParamsSource(nameof(Cases))]
    public CsvParallelInput Input { get; set; }
    [ParamsSource(nameof(Degrees))]
    public int Dop { get; set; }
    public IEnumerable<CsvParallelInput> Cases() => CsvParallelFixture.Cases();
    public IEnumerable<int> Degrees() => CsvParallelFixture.Degrees();

    [GlobalSetup]
    public async Task SetupAsync() {
        CsvBenchmarkFixture.DescribeEngines();
        _fixture = new(Input);
        try {
            await QualifyAsync(ExcelReaderAggregate);
            await QualifyAsync(() => Task.FromResult(OfficeIMOTextOwnedAggregate()));
            await QualifyAsync(OfficeIMOAsyncPathAggregate);
            await QualifyAsync(() => ReadOfficeStreamAsync(probe: true));
        } catch { Cleanup(); throw; }
        Console.WriteLine($"Qualified direct CSV reduction: {Input}, DOP={Dop}, bytes={_fixture.Bytes}, "
            + $"SHA256={_fixture.Fingerprint}, aggregate={_fixture.ExpectedSum}; every consumed field and input-row multiplicity verified. "
            + "Text case reads/decodes the complete file inside timing; async cases read incrementally. "
            + "OfficeIMO uses decoded UTF16 fields, peer uses UTF8 fields; parser allocation is included.");
    }

    private async Task QualifyAsync(Func<Task<long>> operation) {
        var proof = new CsvAggregateProof(_fixture);
        _proof = _peerProof = proof;
        try { await operation(); proof.Complete(); }
        finally { _proof = _peerProof = null; }
    }

    [GlobalCleanup]
    public void Cleanup() => _fixture?.Dispose();

    [Benchmark(Baseline = true)]
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

    [Benchmark]
    public long OfficeIMOTextOwnedAggregate() {
        string text = File.ReadAllText(_fixture.Path);
        AggregateState state = CsvDocument.AggregateTextRowsAsParallel(text, NewState, BuildAccumulator, Merge,
            loadOptions: LoadOptions(), parallelOptions: ParallelOptions());
        return _fixture.Check(state.Sum, state.Count);
    }

    [Benchmark]
    public async Task<long> OfficeIMOAsyncPathAggregate() {
        AggregateState state = await CsvDocument.AggregateRowsAsParallelAsync(_fixture.Path,
            NewState, BuildAccumulator, Merge, loadOptions: LoadOptions(), parallelOptions: ParallelOptions());
        return _fixture.Check(state.Sum, state.Count);
    }

    [Benchmark]
    public Task<long> OfficeIMOAsyncStreamAggregate() => ReadOfficeStreamAsync(probe: false);

    private async Task<long> ReadOfficeStreamAsync(bool probe) {
        await using var file = new FileStream(_fixture.Path, FileMode.Open, FileAccess.Read, FileShare.Read,
            4096, FileOptions.Asynchronous | FileOptions.SequentialScan);
        if (!probe) {
            AggregateState state = await CsvDocument.AggregateRowsAsParallelAsync(file,
                NewState, BuildAccumulator, Merge, loadOptions: LoadOptions(), parallelOptions: ParallelOptions());
            return _fixture.Check(state.Sum, state.Count);
        }
        using var observed = new CsvAggregateAsyncProbe(file);
        AggregateState proved = await CsvDocument.AggregateRowsAsParallelAsync(observed,
            NewState, BuildAccumulator, Merge, loadOptions: LoadOptions(), parallelOptions: ParallelOptions());
        if (observed.AsyncReads == 0 || observed.SyncReads != 0)
            throw new InvalidDataException("CSV async aggregation did not use actual asynchronous input reads.");
        Console.WriteLine($"CSV stream qualification: asyncReads={observed.AsyncReads}; syncReads={observed.SyncReads}.");
        return _fixture.Check(proved.Sum, proved.Count);
    }

    private CsvLoadOptions LoadOptions() => new() {
        MaxInputBytes = Math.Max(CsvLoadOptions.DefaultMaxInputBytes, _fixture.Bytes),
    };
    private ParallelRowMappingOptions ParallelOptions() => new() { MaxDegreeOfParallelism = Dop == 0 ? null : Dop };
    private static AggregateState NewState() => default;
    private static AggregateState Merge(AggregateState result, AggregateState following) =>
        new() { Sum = result.Sum + following.Sum, Count = result.Count + following.Count };

    private CsvRecordAccumulator<AggregateState> BuildAccumulator(CsvRecordHeader header) {
        if (_fixture.Heavy) {
            int region = header.GetOrdinal("Region"), country = header.GetOrdinal("Country"),
                date = header.GetOrdinal("OrderDate"), price = header.GetOrdinal("UnitPrice"),
                revenue = header.GetOrdinal("TotalRevenue"), units = header.GetOrdinal("Units");
            return (ref AggregateState state, CsvRecord row) => {
                ReadOnlySpan<char> regionText = row.GetSpan(region), countryText = row.GetSpan(country);
                DateTime orderDate = row.GetDateTime(date);
                decimal unitPrice = row.GetDecimal(price), totalRevenue = row.GetDecimal(revenue);
                int count = row.GetInt32(units);
                _proof?.Heavy(regionText, countryText, orderDate, unitPrice, totalRevenue, count);
                state.Sum += count; state.Count++;
            };
        }
        int a = header.GetOrdinal("A"), b = header.GetOrdinal("B"), c = header.GetOrdinal("C");
        return (ref AggregateState state, CsvRecord row) => {
            int first = row.GetInt32(a), second = row.GetInt32(b), third = row.GetInt32(c);
            _proof?.Narrow(first, second, third);
            state.Sum += first; state.Count++;
        };
    }

    private struct AggregateState { internal long Sum; internal int Count; }
    private sealed class HeavyAccumulator : ICsvAccumulator<HeavyAccumulator, CsvHeavyRefRecord> {
        private readonly CsvAggregateProof? _validation = _peerProof;
        public long Sum { get; private set; }
        public int Count { get; private set; }
        public void Add(CsvHeavyRefRecord row) {
            if (_validation is not null) _validation.Heavy(Encoding.UTF8.GetString(row.Region),
                Encoding.UTF8.GetString(row.Country), row.OrderDate, row.UnitPrice, row.TotalRevenue, row.Units);
            Sum += row.Units; Count++;
        }
        public void Merge(HeavyAccumulator following) { Sum += following.Sum; Count += following.Count; }
    }
    private sealed class NarrowAccumulator : ICsvAccumulator<NarrowAccumulator, CsvNarrowStructRecord> {
        private readonly CsvAggregateProof? _validation = _peerProof;
        public long Sum { get; private set; }
        public int Count { get; private set; }
        public void Add(CsvNarrowStructRecord row) {
            _validation?.Narrow(row.A, row.B, row.C);
            Sum += row.A; Count++;
        }
        public void Merge(NarrowAccumulator following) { Sum += following.Sum; Count += following.Count; }
    }
}

/// <summary>Untimed source-value and multiplicity proof, independent of partition grouping.</summary>
internal sealed class CsvAggregateProof(CsvParallelFixture fixture) {
    // LCM of the complete heavy generator pattern: price9000/date3650;
    // region4/country6/units500 divide that period.
    private const int HeavyPeriod = 657_000;
    private readonly int[] _seen = new int[Math.Min(fixture.Input.Rows, fixture.Heavy ? HeavyPeriod : fixture.Input.Rows)];
    internal void Heavy(ReadOnlySpan<char> region, ReadOnlySpan<char> country, DateTime date,
        decimal price, decimal revenue, int units) {
        int representative = fixture.ValidateHeavyParts(region, country, date, price, revenue, units);
        Interlocked.Increment(ref _seen[representative]);
    }
    internal void Narrow(int a, int b, int c) {
        if (a < 0 || a >= fixture.Input.Rows || b != a * 3 || c != a * 7)
            throw new InvalidDataException("Direct CSV aggregate integer fields differ.");
        Interlocked.Increment(ref _seen[a]);
    }
    internal void Complete() {
        for (int index = 0; index < _seen.Length; index++) {
            int expected = fixture.Heavy ? (fixture.Input.Rows - 1 - index) / HeavyPeriod + 1 : 1;
            if (_seen[index] != expected) throw new InvalidDataException("Direct CSV aggregate lost or duplicated input records.");
        }
    }
}

/// <summary>Observes input I/O only during qualification; measured streams remain ordinary FileStreams.</summary>
internal sealed class CsvAggregateAsyncProbe(Stream source) : Stream {
    internal int AsyncReads { get; private set; }
    internal int SyncReads { get; private set; }
    public override bool CanRead => source.CanRead;
    public override bool CanSeek => source.CanSeek;
    public override bool CanWrite => false;
    public override long Length => source.Length;
    public override long Position { get => source.Position; set => source.Position = value; }
    public override int Read(byte[] buffer, int offset, int count) { SyncReads++; return source.Read(buffer, offset, count); }
    public override int Read(Span<byte> buffer) { SyncReads++; return source.Read(buffer); }
    public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken token = default) {
        AsyncReads++; return source.ReadAsync(buffer, token);
    }
    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token) {
        AsyncReads++; return source.ReadAsync(buffer, offset, count, token);
    }
    public override long Seek(long offset, SeekOrigin origin) => source.Seek(offset, origin);
    public override void Flush() => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
}
#endif
