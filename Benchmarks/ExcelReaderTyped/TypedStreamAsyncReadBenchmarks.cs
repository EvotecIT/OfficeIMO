#if OFFICEIMO_BENCHMARK_NEW_APIS
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader.Xlsx;
using OfficeIMO.Data;
using Sylvan.Data;
using Sylvan.Data.Excel;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Adds asynchronous stream opening while preserving the existing byte-open async lane.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("TypedStreamOpenAsync")]
    public class TypedStreamAsyncReadBenchmarks {
        private readonly TypedReadBenchmarks _fixture = new();
        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 50_000;
        [ParamsSource(nameof(Shapes))]
        public string Shape { get; set; } = "Original";
        public IEnumerable<int> RowCounts() => _fixture.RowCounts();
        public IEnumerable<string> Shapes() => _fixture.Shapes();
        [GlobalSetup]
        public async Task SetupAsync() {
            _fixture.RowCount = RowCount;
            _fixture.Shape = Shape;
            await _fixture.SetupAsync();
            await using MemoryStream stream = new MemoryStream(_fixture.Workbook, writable: false);
            await using ExcelWorkbookDataReader reader = await ExcelDocument.OpenDataReaderAsync(stream, new ExcelReadOptions { HasHeaderRow = true });
            int count = 0;
            await foreach (TypedRecord record in reader.RowsAsAsync<TypedRecord>())
                TypedWorkbookFixture.ValidateRecord(record, ++count);
            if (count != RowCount) throw new InvalidDataException("Async stream projection row count differs.");
            await ValidatePeersAsync();
            await OfficeIMOStreamOpenTypedAsync();
            await ExcelReaderStreamOpenTypedAsync();
            await SylvanStreamOpenTypedAsync();
            Console.WriteLine($"Qualified public async stream typed opening: shape={Shape}, rows={RowCount}.");
        }
        [Benchmark]
        public async Task<long> OfficeIMOStreamOpenTypedAsync() {
            await using MemoryStream stream = new MemoryStream(_fixture.Workbook, writable: false);
            await using ExcelWorkbookDataReader reader = await ExcelDocument.OpenDataReaderAsync(stream, new ExcelReadOptions { HasHeaderRow = true });
            int count = 0;
            long sum = 0;
            await foreach (TypedRecord record in reader.RowsAsAsync<TypedRecord>()) {
                sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record));
                count++;
            }
            return _fixture.Check(sum, count);
        }
        [Benchmark(Baseline = true)]
        public async Task<long> ExcelReaderStreamOpenTypedAsync() {
            await using MemoryStream stream = new MemoryStream(_fixture.Workbook, writable: false);
            await using XlsxWorkbook workbook = await ExcelReaderApi.FromXlsxAsync(stream);
            long sum = 0;
            int count = 0;
            await foreach (TypedRecord record in ExcelParser.FromAttributes<TypedRecord>().Parse(workbook.FirstSheet)) {
                sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record));
                count++;
            }
            return _fixture.Check(sum, count);
        }
        [Benchmark]
        public async Task<long> SylvanStreamOpenTypedAsync() {
            await using MemoryStream stream = new MemoryStream(_fixture.Workbook, writable: false);
            await using ExcelDataReader reader = await global::Sylvan.Data.Excel.ExcelDataReader.CreateAsync(stream,
                ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
            long sum = 0;
            int count = 0;
            await foreach (TypedRecord record in reader.GetRecordsAsync<TypedRecord>()) {
                sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record));
                count++;
            }
            return _fixture.Check(sum, count);
        }
        private async Task ValidatePeersAsync() {
            await using MemoryStream stream = new MemoryStream(_fixture.Workbook, writable: false);
            await using XlsxWorkbook workbook = await ExcelReaderApi.FromXlsxAsync(stream);
            int count = 0;
            await foreach (TypedRecord record in ExcelParser.FromAttributes<TypedRecord>().Parse(workbook.FirstSheet))
                TypedWorkbookFixture.ValidateRecord(record, ++count);
            if (count != RowCount) throw new InvalidDataException("Peer async stream row count differs.");
            await using MemoryStream sylvanStream = new MemoryStream(_fixture.Workbook, writable: false);
            await using ExcelDataReader reader = await global::Sylvan.Data.Excel.ExcelDataReader.CreateAsync(sylvanStream,
                ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions());
            count = 0;
            await foreach (TypedRecord record in reader.GetRecordsAsync<TypedRecord>())
                TypedWorkbookFixture.ValidateRecord(record, ++count);
            if (count != RowCount) throw new InvalidDataException("Sylvan async stream row count differs.");
        }
    }
}
#endif