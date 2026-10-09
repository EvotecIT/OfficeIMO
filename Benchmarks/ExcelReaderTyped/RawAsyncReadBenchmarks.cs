using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Consumes identical raw workbooks through each library's public async reader APIs.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("RawAsyncOriginal")]
    public class RawAsyncReadBenchmarks {
        private readonly RawReadBenchmarks _workload = new();

        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 50_000;

        [ParamsAllValues]
        public RawWorkbookFormat Format { get; set; }

        public IEnumerable<int> RowCounts() => _workload.RowCounts();

        [GlobalSetup]
        public async Task SetupAsync() {
            _workload.RowCount = RowCount;
            _workload.Format = Format;
            await _workload.SetupAsync();
            await _workload.ValidateAsyncReaders();
        }

        [Benchmark(Baseline = true)]
        public Task<long> ExcelReaderAsync() => _workload.ExcelReaderAsync(materializeStrings: false);
        [Benchmark]
        public Task<long> OfficeIMOAsync() => _workload.OfficeIMOAsync();
        [Benchmark]
        public Task<long> SylvanAsync() => _workload.SylvanAsync();
    }

    /// <summary>Compares the same raw async APIs with materialized strings in every engine.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("RawAsyncStringsMaterialized")]
    public class MaterializedRawAsyncReadBenchmarks {
        private readonly RawReadBenchmarks _workload = new();

        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 50_000;

        [ParamsAllValues]
        public RawWorkbookFormat Format { get; set; }

        public IEnumerable<int> RowCounts() => _workload.RowCounts();

        [GlobalSetup]
        public async Task SetupAsync() {
            _workload.RowCount = RowCount;
            _workload.Format = Format;
            await _workload.SetupAsync();
            await _workload.ValidateAsyncReaders();
        }

        [Benchmark(Baseline = true)]
        public Task<long> ExcelReaderStringsMaterializedAsync() => _workload.ExcelReaderAsync(materializeStrings: true);
        [Benchmark]
        public Task<long> OfficeIMOAsync() => _workload.OfficeIMOAsync();
        [Benchmark]
        public Task<long> SylvanAsync() => _workload.SylvanAsync();
    }
}
