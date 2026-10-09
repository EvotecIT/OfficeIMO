namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    internal static class CsvBenchmarkQualification {
#if OFFICEIMO_BENCHMARK_NEW_APIS
        internal static async Task ValidateDirectAggregateScaleAsync() {
            // An explicit scale command preserves the bounded ordinary qualification.
            foreach (CsvParallelInput input in CsvParallelFixture.Cases())
                foreach (int degree in CsvParallelFixture.Degrees()) {
                    CsvDirectAggregateBenchmarks workload = new CsvDirectAggregateBenchmarks { Input = input, Dop = degree };
                    try { await workload.SetupAsync(); }
                    finally { workload.Cleanup(); }
                    Console.WriteLine($"CSV direct aggregate scale qualification passed for {input}, DOP={degree}.");
                }
        }

        internal static async Task ValidateDirectAggregateAsync() {
            foreach (string corpus in new[] { "ConversionHeavy", "NarrowInt", "ConversionHeavy3M" })
                foreach (int degree in CsvParallelFixture.Degrees()) {
                    int rows = int.Parse(Environment.GetEnvironmentVariable("OFFICEIMO_CSV_PARALLEL_BENCHMARK_ROWS") ?? "50000",
                        System.Globalization.CultureInfo.InvariantCulture);
                    CsvDirectAggregateBenchmarks workload = new CsvDirectAggregateBenchmarks { Input = new(corpus, rows), Dop = degree };
                    try { await workload.SetupAsync(); }
                    finally { workload.Cleanup(); }
                }
        }

        internal static void ValidateUtf8Writer() {
            CsvUtf8WriterBenchmarks workload = new CsvUtf8WriterBenchmarks();
            try { workload.Setup(); }
            finally { workload.Cleanup(); }
        }
#endif
        // This command qualifies parallel contracts with bounded fixtures. The BDN
        // parameter source retains the original multi-million row counts by default.
        internal static async Task ValidateParallelAsync() {
#if !OFFICEIMO_BENCHMARK_NEW_APIS
            Console.WriteLine("Decoded-text default-date CSV lanes require current source: published3.4.4 crashes on batch dates at DOP1; its Core parallel lane remains qualified.");
#endif
            foreach (string corpus in new[] { "ConversionHeavy", "NarrowInt", "ConversionHeavy3M" }) {
                foreach (int degree in CsvParallelFixture.Degrees()) {
                    int rows = int.Parse(Environment.GetEnvironmentVariable("OFFICEIMO_CSV_PARALLEL_BENCHMARK_ROWS") ?? "50000",
                        System.Globalization.CultureInfo.InvariantCulture);
                    CsvParallelTypedBenchmarks typed = new CsvParallelTypedBenchmarks { Input = new(corpus, rows), Dop = degree };
                    try { await typed.SetupAsync(); }
                    finally { typed.Cleanup(); }
                    CsvParallelAggregateBenchmarks aggregate = new CsvParallelAggregateBenchmarks { Input = new(corpus, rows), Dop = degree };
                    try { await aggregate.SetupAsync(); }
                    finally { aggregate.Cleanup(); }
                    Console.WriteLine($"CSV parallel typed/aggregate qualification passed for {corpus}, rows={rows}, DOP={degree}.");
                }
            }
        }

        internal static async Task ValidateAsync() {
            foreach (int rows in new TypedReadBenchmarks().RowCounts()) {
#if OFFICEIMO_BENCHMARK_NEW_APIS
                new CsvRawReadBenchmarks { RowCount = rows }.Setup();
#else
                Console.WriteLine("Borrowed CSV original lane requires OfficeIMOBenchmarkNewApis=true and a compatible source/assembly input.");
#endif
                new CsvMaterializedReadBenchmarks { RowCount = rows }.Setup();
                new CsvTypedReadBenchmarks { RowCount = rows }.Setup();
                await new CsvRawAsyncBenchmarks { RowCount = rows }.SetupAsync();
                await new CsvTypedAsyncBenchmarks { RowCount = rows }.SetupAsync();
                new CsvWideReadBenchmarks { RowCount = rows, MaterializeStrings = true }.Setup();
#if OFFICEIMO_BENCHMARK_NEW_APIS
                new CsvWideReadBenchmarks { RowCount = rows, MaterializeStrings = false }.Setup();
#endif
                new CsvWriterBenchmarks { RowCount = rows }.Setup();
                Console.WriteLine($"CSV sequential/async/writer qualification passed for {rows} rows.");
            }
        }
    }
}
