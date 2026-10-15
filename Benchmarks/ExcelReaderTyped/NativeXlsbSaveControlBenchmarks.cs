#if OFFICEIMO_BENCHMARK_NEW_APIS
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>
    /// Measures warmed document creation, explicit object insertion, XLSB save and disposal.
    /// The source records and fixed destination are prepared outside the operation.
    /// </summary>
    [MemoryDiagnoser]
    [SimpleJob(launchCount: 1, warmupCount: 4, iterationCount: 8, invocationCount: 1)]
    [InvocationCount(1, 1)]
    [BenchmarkCategory("NativeXlsbWarmSaveControl")]
    public class NativeXlsbWarmSaveControlBenchmarks {
        private readonly NativeXlsbSaveControlWorkload _workload = new();
        [Params("Mixed", "AllText")]
        public string Shape { get; set; } = "Mixed";
        [Params("Inline", "Shared")]
        public string Storage { get; set; } = "Inline";
        [GlobalSetup]
        public void Setup() {
            _workload.Prepare(Shape, Storage);
            _workload.Save();
        }
        [Benchmark]
        public long OfficeIMOModelAndSave() => _workload.Save();
        [GlobalCleanup]
        public void Cleanup() => _workload.Dispose();
    }

    /// <summary>
    /// Measures the first complete XLSB construction and save in each fresh process.
    /// Setup prepares records and a destination without calling a writer or output reader.
    /// </summary>
    // MemoryDiagnoser runs an extra workload for statistics, so this control is time-only.
    [SimpleJob(RunStrategy.ColdStart, launchCount: 16, warmupCount: 0, iterationCount: 1, invocationCount: 1)]
    [InvocationCount(1, 1)]
    [BenchmarkCategory("NativeXlsbFirstSaveControl")]
    public class NativeXlsbFirstSaveControlBenchmarks {
        private readonly NativeXlsbSaveControlWorkload _workload = new();
        [Params("Mixed", "AllText")]
        public string Shape { get; set; } = "Mixed";
        [Params("Inline", "Shared")]
        public string Storage { get; set; } = "Inline";
        [GlobalSetup]
        public void Setup() => _workload.Prepare(Shape, Storage);
        [Benchmark]
        public long OfficeIMOModelAndSave() => _workload.Save();
        [GlobalCleanup]
        public void Cleanup() => _workload.Dispose();
    }
}
#endif
