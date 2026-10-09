#if OFFICEIMO_BENCHMARK_ARROW && OFFICEIMO_BENCHMARK_NEW_APIS
using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Exact upstream one-batch Arrow conversion with matched explicit schemas.</summary>
    [MemoryDiagnoser]
    public class ArrowConversionBenchmarks {
        private ArrowConversionWorkload _workload = null!;
        [Params(100_000)] public int RowCount { get; set; }
        [ParamsAllValues] public ArrowConversionScenario Scenario { get; set; }

        [GlobalSetup]
        public async Task SetupAsync() {
            BenchmarkInput.WriteDescription();
            _workload = await ArrowConversionWorkload.CreateAsync(RowCount);
            Validate();
        }

        internal void Validate() {
            Action<Apache.Arrow.RecordBatch, int> validate = (batch, offset) => {
                if (offset != 0 || batch.Length != RowCount) throw new InvalidDataException("Expected one complete Arrow batch.");
                ArrowConversionValidation.ValidateExplicit(batch, offset, Scenario, RowCount);
            };
            if (_workload.ConvertPeer(Scenario, validate) != RowCount
                || _workload.ConvertOfficeIMO(Scenario, RowCount, validate) != RowCount)
                throw new InvalidDataException("Arrow conversion row count differs.");
        }

        [Benchmark(Baseline = true)] public long ExcelReader() => _workload.ConvertPeer(Scenario);
        [Benchmark] public long OfficeIMO() => _workload.ConvertOfficeIMO(Scenario, RowCount);
    }

    /// <summary>OfficeIMO's bounded-batch contract, measured separately from one-batch conversion.</summary>
    [MemoryDiagnoser]
    public class ArrowBoundedConversionBenchmarks {
        private const int BatchSize = 8_192;
        private ArrowConversionWorkload _workload = null!;
        [Params(100_000)] public int RowCount { get; set; }
        [ParamsAllValues] public ArrowConversionScenario Scenario { get; set; }

        [GlobalSetup]
        public async Task SetupAsync() {
            BenchmarkInput.WriteDescription();
            _workload = await ArrowConversionWorkload.CreateAsync(RowCount);
            if (_workload.ConvertOfficeIMO(Scenario, BatchSize,
                (batch, offset) => ArrowConversionValidation.ValidateExplicit(batch, offset, Scenario, RowCount)) != RowCount)
                throw new InvalidDataException("Bounded Arrow row count differs.");
        }

        [Benchmark] public long OfficeIMO_BoundedBatches() => _workload.ConvertOfficeIMO(Scenario, BatchSize);
    }

    /// <summary>
    /// The upstream inference method produces string columns because parseText defaults to false.
    /// OfficeIMO's public CSV inference produces typed columns and is validated but not paired for timing.
    /// </summary>
    [MemoryDiagnoser]
    public class ArrowInferredConversionBenchmarks {
        private ArrowConversionWorkload _workload = null!;
        [Params(100_000)] public int RowCount { get; set; }

        [GlobalSetup]
        public async Task SetupAsync() {
            BenchmarkInput.WriteDescription();
            _workload = await ArrowConversionWorkload.CreateAsync(RowCount);
            if (_workload.ConvertPeerInferred((batch, offset) => ArrowConversionValidation.ValidatePeerInference(batch, offset, RowCount)) != RowCount
                || _workload.ConvertOfficeIMOInferred((batch, offset) => ArrowConversionValidation.ValidateOfficeInference(batch, offset, RowCount)) != RowCount)
                throw new InvalidDataException("Inferred Arrow row count differs.");
            Console.WriteLine("Inferred Arrow schemas differ: ExcelReader has four strings; OfficeIMO has string/int32/timestamp/decimal. No paired inference ratio.");
        }

        [Benchmark] public long ExcelReader_CsvTypedWithInference() => _workload.ConvertPeerInferred();
    }

    public static class ArrowBenchmarkValidation {
        public static async Task RunAsync() {
            foreach (ArrowConversionScenario scenario in Enum.GetValues<ArrowConversionScenario>()) {
                await new ArrowConversionBenchmarks { RowCount = 100_000, Scenario = scenario }.SetupAsync();
                await new ArrowBoundedConversionBenchmarks { RowCount = 100_000, Scenario = scenario }.SetupAsync();
                Console.WriteLine($"Arrow {scenario}: required schema, every value and row order validated; one batch and bounded batches.");
            }
            await new ArrowInferredConversionBenchmarks { RowCount = 100_000 }.SetupAsync();
        }
    }
}
#endif