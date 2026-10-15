using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader.Csv;
using OfficeIMO.CSV;
using OfficeIMO.Data;
using System.Data.Common;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Maps the headered four-field CSV into the same materialized value record.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("CsvTypedStruct")]
    public class CsvTypedStructReadBenchmarks {
        private byte[] _data = [];
        private long _expected;

        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 50_000;
        public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();

        [GlobalSetup]
        public void Setup() {
            _data = CsvBenchmarkFixture.Create(RowCount, headers: true);
            _expected = TypedWorkbookFixture.ExpectedChecksum(RowCount);
            CsvBenchmarkFixture.Describe(_data, RowCount, "typed struct headered");
            foreach (int engine in new[] { 0, 1 }) {
                int count = 0;
                long sum = 0;
                foreach (TypedValueRecord record in Records(engine)) {
                    // Reuse the complete native-field proof outside the timed methods.
                    TypedWorkbookFixture.ValidateRecord(new TypedRecord {
                        Name = record.Name, Id = record.Id, Date = record.Date, Value = record.Value,
                    }, ++count);
                    sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record));
                }
                CsvBenchmarkFixture.Check(sum, count, RowCount, _expected);
            }
            using (MemoryStream stream = new MemoryStream(_data, false))
            using (DbDataReader reader = CsvDocument.OpenDataReader(stream))
                CsvBenchmarkFixture.Validate(reader, RowCount, headers: true);
            ExcelReaderTyped();
            OfficeIMORowsAs();
            Console.WriteLine($"Validated typed CSV struct mapping: rows={RowCount}, all four native fields, headers and row order.");
        }

        [Benchmark(Baseline = true)]
        public long ExcelReaderTyped() {
            using MemoryStream stream = new MemoryStream(_data, false);
            using CsvReader workbook = ExcelReaderApi.FromCsv(stream);
            long sum = 0;
            int count = 0;
            // Keep the concrete CsvSheet overload, including its CSV text-date converter.
            foreach (TypedValueRecord record in ExcelParser.FromAttributes<TypedValueRecord>().Parse(workbook.FirstSheet)) {
                sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record));
                count++;
            }
            return CsvBenchmarkFixture.Check(sum, count, RowCount, _expected);
        }

        [Benchmark]
        public long OfficeIMORowsAs() {
            using MemoryStream stream = new MemoryStream(_data, false);
            using DbDataReader reader = CsvDocument.OpenDataReader(stream);
            long sum = 0;
            int count = 0;
            foreach (TypedValueRecord record in reader.RowsAs<TypedValueRecord>()) {
                sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record));
                count++;
            }
            return CsvBenchmarkFixture.Check(sum, count, RowCount, _expected);
        }

        private IEnumerable<TypedValueRecord> Records(int engine) {
            using MemoryStream stream = new MemoryStream(_data, false);
            if (engine == 0) {
                using CsvReader workbook = ExcelReaderApi.FromCsv(stream);
                foreach (TypedValueRecord record in ExcelParser.FromAttributes<TypedValueRecord>().Parse(workbook.FirstSheet))
                    yield return record;
            } else {
                using DbDataReader reader = CsvDocument.OpenDataReader(stream);
                foreach (TypedValueRecord record in reader.RowsAs<TypedValueRecord>()) yield return record;
            }
        }
    }
}
