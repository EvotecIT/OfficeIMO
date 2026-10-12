using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader.Xlsb;
using OfficeIMO.Data;
using Sylvan.Data;
using Sylvan.Data.Excel;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Materializes the same XLSB records through public async enumeration APIs.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("TypedXlsbAsync")]
    public class TypedXlsbAsyncReadBenchmarks {
        private readonly TypedXlsbReadBenchmarks _workload = new();

        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 50_000;

        public IEnumerable<int> RowCounts() => _workload.RowCounts();

        [GlobalSetup]
        public async Task SetupAsync() {
            _workload.RowCount = RowCount;
            await _workload.SetupAsync();
            await ValidateAsync(ReadOfficeIMOForValidation());
            await ValidateAsync(ReadExcelReaderForValidation());
            await ValidateAsync(ReadSylvanForValidation());
            await ExcelReaderTypedXlsbAsync();
            await OfficeIMOTypedXlsbAsync();
            await SylvanTypedXlsbAsync();
            Console.WriteLine($"Validated async typed XLSB: rows={RowCount}; all fields checked.");
        }

        [Benchmark(Baseline = true)]
        public async Task<long> ExcelReaderTypedXlsbAsync() {
            await using MemoryStream stream = new MemoryStream(_workload.Workbook, writable: false);
            await using XlsbWorkbook workbook = await ExcelReaderApi.FromXlsbAsync(stream);
            long sum = 0;
            int count = 0;
            await foreach (TypedRecord record in ExcelParser.FromAttributes<TypedRecord>().Parse(workbook.FirstSheet)) {
                sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record));
                count++;
            }
            return _workload.Check(sum, count);
        }

        [Benchmark]
        public async Task<long> OfficeIMOTypedXlsbAsync() {
            await using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(_workload.Workbook, new ExcelReadOptions { HasHeaderRow = true });
            long sum = 0;
            int count = 0;
            await foreach (TypedRecord record in reader.RowsAsAsync<TypedRecord>()) {
                sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record));
                count++;
            }
            return _workload.Check(sum, count);
        }

        [Benchmark]
        public async Task<long> SylvanTypedXlsbAsync() {
            await using MemoryStream stream = new MemoryStream(_workload.Workbook, writable: false);
            await using ExcelDataReader reader = await global::Sylvan.Data.Excel.ExcelDataReader.CreateAsync(stream,
                ExcelWorkbookType.ExcelBinary, new ExcelDataReaderOptions());
            long sum = 0;
            int count = 0;
            await foreach (TypedRecord record in reader.GetRecordsAsync<TypedRecord>()) {
                sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record));
                count++;
            }
            return _workload.Check(sum, count);
        }

        private async Task ValidateAsync(IAsyncEnumerable<TypedRecord> records) {
            int count = 0;
            await foreach (TypedRecord record in records) TypedWorkbookFixture.ValidateRecord(record, ++count);
            if (count != RowCount) throw new InvalidDataException("Incorrect async typed XLSB row count.");
        }

        private async IAsyncEnumerable<TypedRecord> ReadOfficeIMOForValidation() {
            await using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(_workload.Workbook, new ExcelReadOptions { HasHeaderRow = true });
            await foreach (TypedRecord record in reader.RowsAsAsync<TypedRecord>()) yield return record;
        }

        private async IAsyncEnumerable<TypedRecord> ReadExcelReaderForValidation() {
            await using MemoryStream stream = new MemoryStream(_workload.Workbook, writable: false);
            await using XlsbWorkbook workbook = await ExcelReaderApi.FromXlsbAsync(stream);
            await foreach (TypedRecord record in ExcelParser.FromAttributes<TypedRecord>().Parse(workbook.FirstSheet)) yield return record;
        }

        private async IAsyncEnumerable<TypedRecord> ReadSylvanForValidation() {
            await using MemoryStream stream = new MemoryStream(_workload.Workbook, writable: false);
            await using ExcelDataReader reader = await global::Sylvan.Data.Excel.ExcelDataReader.CreateAsync(stream,
                ExcelWorkbookType.ExcelBinary, new ExcelDataReaderOptions());
            await foreach (TypedRecord record in reader.GetRecordsAsync<TypedRecord>()) yield return record;
        }
    }
}
