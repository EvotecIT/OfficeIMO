using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using OfficeIMO.CSV;
using Sylvan.Data.Csv;
using System.Data.Common;
using System.Globalization;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Async headerless scalar reads with materialized names for every engine.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("CsvRawAsyncMaterialized")]
    public class CsvRawAsyncBenchmarks {
        private byte[] _data = [];
        private long _expected;
        private bool _validate;
        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 50_000;
        public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();

        [GlobalSetup]
        public async Task SetupAsync() {
            _data = CsvBenchmarkFixture.Create(RowCount, headers: false);
            _expected = TypedWorkbookFixture.ExpectedChecksum(RowCount);
            CsvBenchmarkFixture.Describe(_data, RowCount, "raw async materialized name");
            _validate = true;
            await ExcelReaderAsync(); await OfficeIMOAsync(); await SylvanAsync();
            _validate = false;
        }

        [Benchmark(Baseline = true)]
        public async Task<long> ExcelReaderAsync() {
            await using MemoryStream stream = new MemoryStream(_data, false);
            await using CsvReader workbook = ExcelReaderApi.FromCsv(stream);
            await using CsvReader.Enumerator rows = workbook.FirstSheet.GetAsyncEnumerator();
            int count = 0;
            long sum = 0;
            while (await rows.MoveNextAsync()) {
                Row row = rows.Current;
                if (!row[1].TryParse(CultureInfo.InvariantCulture, out int id)
                    || !row[2].TryGetDateTime(out DateTime date)
                    || !row[3].TryParse(CultureInfo.InvariantCulture, out double value))
                    throw new InvalidDataException("CSV scalar parsing failed.");
                string name = row[0].GetString();
                if (_validate) {
                    if (row.ColumnCount != 4) throw new InvalidDataException("CSV field count differs.");
                    TypedWorkbookFixture.ValidateRecord(new TypedRecord { Name = name, Id = id, Date = date, Value = value }, count + 1);
                }
                sum = unchecked(sum + name.Length + id + date.Ticks + (long)value);
                count++;
            }
            return CsvBenchmarkFixture.Check(sum, count, RowCount, _expected);
        }

        [Benchmark]
        public async Task<long> OfficeIMOAsync() {
            await using MemoryStream stream = new MemoryStream(_data, false);
            using DbDataReader reader = await CsvDocument.OpenDataReaderAsync(stream, new CsvLoadOptions { HasHeaderRow = false });
            return await ReadAsync(reader);
        }

        [Benchmark]
        public async Task<long> SylvanAsync() {
            await using MemoryStream stream = new MemoryStream(_data, false);
            using StreamReader text = new StreamReader(stream);
            using CsvDataReader reader = await global::Sylvan.Data.Csv.CsvDataReader.CreateAsync(text,
                new global::Sylvan.Data.Csv.CsvDataReaderOptions { HasHeaders = false, Culture = CultureInfo.InvariantCulture });
            return await ReadAsync(reader);
        }

        private async Task<long> ReadAsync(DbDataReader reader) {
            int count = 0;
            long sum = 0;
            while (await reader.ReadAsync()) {
                string name = reader.GetString(0);
                int id = reader.GetInt32(1);
                DateTime date = reader.GetDateTime(2);
                double value = reader.GetDouble(3);
                if (_validate) {
                    if (reader.FieldCount != 4) throw new InvalidDataException("CSV field count differs.");
                    TypedWorkbookFixture.ValidateRecord(new TypedRecord { Name = name, Id = id, Date = date, Value = value }, count + 1);
                }
                sum = unchecked(sum + name.Length + id + date.Ticks + (long)value);
                count++;
            }
            return CsvBenchmarkFixture.Check(sum, count, RowCount, _expected);
        }
    }
}
