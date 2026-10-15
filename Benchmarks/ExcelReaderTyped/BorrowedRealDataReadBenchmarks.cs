#if OFFICEIMO_BENCHMARK_NEW_APIS
using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Consumes public UTF8 text and native fields from the authentic pinned real-data workbooks.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("RealDataPublicUtf8")]
    public class BorrowedRealDataReadBenchmarks {
        private readonly RealDataReadBenchmarks _workload = new();

        [Params(ComparisonWorkbookFormat.Xlsx, ComparisonWorkbookFormat.Xlsm, ComparisonWorkbookFormat.Xlsb)]
        public ComparisonWorkbookFormat Format { get; set; }

        [GlobalSetup]
        public void Setup() {
            _workload.Format = Format;
            _workload.Setup();
            _workload.QualifyBorrowedUtf8();
        }

        [Benchmark(Baseline = true)]
        public long ExcelReaderBorrowed() => _workload.ExcelReaderOriginal();
        [Benchmark]
        public long ExcelReaderBorrowedPrefetch() => _workload.ExcelReaderPrefetch(memory: false);
        [Benchmark]
        public long OfficeIMOBorrowedUtf8() => _workload.OfficeIMOBorrowedUtf8();
    }
}
#endif
