#if OFFICEIMO_BENCHMARK_NEW_APIS
using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Consumes equivalent public UTF8 text and native fields from the generated raw workbooks.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("RawPublicUtf8")]
    public class BorrowedRawReadBenchmarks {
        private readonly WorkbookScanWorkload _workload = new();

        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 50_000;
        [Params(RawWorkbookFormat.Xlsx, RawWorkbookFormat.Xlsb)]
        public RawWorkbookFormat Format { get; set; }
        public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();

        [GlobalSetup]
        public async Task SetupAsync() {
            BenchmarkInput.WriteDescription();
            byte[] bytes = await RawWorkbookFixture.CreateAsync(RowCount, Format);
            _workload.Setup(bytes, Format switch {
                RawWorkbookFormat.Xlsx => ComparisonWorkbookFormat.Xlsx,
                RawWorkbookFormat.Xlsb => ComparisonWorkbookFormat.Xlsb,
                _ => throw new ArgumentOutOfRangeException(nameof(Format)),
            }, RowCount, 4, index => {
                TypedRecord expected = TypedWorkbookFixture.ExpectedRecord(index + 1);
                return [expected.Name!, (double)expected.Id, expected.Date, expected.Value];
            });
            _workload.QualifyBorrowedUtf8();
        }

        [Benchmark(Baseline = true)]
        public long ExcelReaderBorrowed() => _workload.ExcelReader();
        [Benchmark]
        public long ExcelReaderBorrowedPrefetch() => _workload.ExcelReader(prefetch: true);
        [Benchmark]
        public long OfficeIMOBorrowedUtf8() => _workload.OfficeIMOBorrowedUtf8();
    }
}
#endif
