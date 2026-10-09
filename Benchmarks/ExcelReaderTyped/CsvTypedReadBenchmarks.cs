using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader.Csv;
using OfficeIMO.CSV;
using OfficeIMO.Data;
using Sylvan.Data;
using System.Data.Common;
using System.Globalization;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;
using SylvanCsv = Sylvan.Data.Csv.CsvDataReader;
using SylvanOptions = Sylvan.Data.Csv.CsvDataReaderOptions;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>All libraries map the headered input into the same materialized record class.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("CsvTyped")]
    public class CsvTypedReadBenchmarks {
        private byte[] _data = [];
        private long _expected;
        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 50_000;
        public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();

        [GlobalSetup]
        public void Setup() {
            _data = CsvBenchmarkFixture.Create(RowCount, headers: true);
            _expected = TypedWorkbookFixture.ExpectedChecksum(RowCount);
            CsvBenchmarkFixture.Describe(_data, RowCount, "typed headered");
            foreach (int engine in new[] { 0, 1, 2 }) {
                int count = 0;
                foreach (TypedRecord record in Records(engine)) TypedWorkbookFixture.ValidateRecord(record, ++count);
                if (count != RowCount) throw new InvalidDataException("Typed CSV row count differs.");
            }
            using (MemoryStream stream = new MemoryStream(_data, false))
            using (DbDataReader reader = CsvDocument.OpenDataReader(stream)) CsvBenchmarkFixture.Validate(reader, RowCount, headers: true);
            ExcelReaderTyped(); OfficeIMORowsAs(); SylvanTyped();
        }

        [Benchmark(Baseline = true)]
        public long ExcelReaderTyped() {
            using MemoryStream stream = new MemoryStream(_data, false);
            using CsvReader workbook = ExcelReaderApi.FromCsv(stream);
            long sum = 0;
            int count = 0;
            foreach (TypedRecord record in ExcelParser.FromAttributes<TypedRecord>().Parse(workbook.FirstSheet)) {
                sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record)); count++;
            }
            return CsvBenchmarkFixture.Check(sum, count, RowCount, _expected);
        }

        [Benchmark]
        public long OfficeIMORowsAs() {
            using MemoryStream stream = new MemoryStream(_data, false);
            using DbDataReader reader = CsvDocument.OpenDataReader(stream);
            long sum = 0;
            int count = 0;
            foreach (TypedRecord record in reader.RowsAs<TypedRecord>()) {
                sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record)); count++;
            }
            return CsvBenchmarkFixture.Check(sum, count, RowCount, _expected);
        }

        [Benchmark]
        public long SylvanTyped() {
            using MemoryStream stream = new MemoryStream(_data, false);
            using StreamReader text = new StreamReader(stream);
            using SylvanCsv reader = SylvanCsv.Create(text, new SylvanOptions { Culture = CultureInfo.InvariantCulture });
            long sum = 0;
            int count = 0;
            foreach (TypedRecord record in reader.GetRecords<TypedRecord>()) {
                sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record)); count++;
            }
            return CsvBenchmarkFixture.Check(sum, count, RowCount, _expected);
        }

        private IEnumerable<TypedRecord> Records(int engine) {
            using MemoryStream stream = new MemoryStream(_data, false);
            if (engine == 0) {
                using CsvReader workbook = ExcelReaderApi.FromCsv(stream);
                foreach (TypedRecord record in ExcelParser.FromAttributes<TypedRecord>().Parse(workbook.FirstSheet)) yield return record;
            } else if (engine == 1) {
                using DbDataReader reader = CsvDocument.OpenDataReader(stream);
                foreach (TypedRecord record in reader.RowsAs<TypedRecord>()) yield return record;
            } else {
                using StreamReader text = new StreamReader(stream);
                using SylvanCsv reader = SylvanCsv.Create(text, new SylvanOptions { Culture = CultureInfo.InvariantCulture });
                foreach (TypedRecord record in reader.GetRecords<TypedRecord>()) yield return record;
            }
        }
    }

    /// <summary>Async source opening, advancement and materialized typed mapping.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("CsvTypedAsync")]
    public class CsvTypedAsyncBenchmarks {
        private byte[] _data = [];
        private long _expected;
        private bool _validate;
        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 50_000;
        public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();

        [GlobalSetup]
        public async Task SetupAsync() {
            _data = CsvBenchmarkFixture.Create(RowCount, headers: true);
            _expected = TypedWorkbookFixture.ExpectedChecksum(RowCount);
            CsvBenchmarkFixture.Describe(_data, RowCount, "typed async headered");
            _validate = true;
            await ExcelReaderTypedAsync(); await OfficeIMORowsAsAsync(); await SylvanTypedAsync();
            _validate = false;
        }

        [Benchmark(Baseline = true)]
        public async Task<long> ExcelReaderTypedAsync() {
            await using MemoryStream stream = new MemoryStream(_data, false);
            await using CsvReader workbook = ExcelReaderApi.FromCsv(stream);
            int count = 0;
            long sum = 0;
            await foreach (TypedRecord record in ExcelParser.FromAttributes<TypedRecord>().Parse(workbook.FirstSheet)) {
                if (_validate) TypedWorkbookFixture.ValidateRecord(record, count + 1);
                sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record)); count++;
            }
            return CsvBenchmarkFixture.Check(sum, count, RowCount, _expected);
        }

        [Benchmark]
        public async Task<long> OfficeIMORowsAsAsync() {
            await using MemoryStream stream = new MemoryStream(_data, false);
            using DbDataReader reader = await CsvDocument.OpenDataReaderAsync(stream);
            int count = 0;
            long sum = 0;
            await foreach (TypedRecord record in reader.RowsAsAsync<TypedRecord>()) {
                if (_validate) TypedWorkbookFixture.ValidateRecord(record, count + 1);
                sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record)); count++;
            }
            return CsvBenchmarkFixture.Check(sum, count, RowCount, _expected);
        }

        [Benchmark]
        public async Task<long> SylvanTypedAsync() {
            await using MemoryStream stream = new MemoryStream(_data, false);
            using StreamReader text = new StreamReader(stream);
            using SylvanCsv reader = await SylvanCsv.CreateAsync(text, new SylvanOptions { Culture = CultureInfo.InvariantCulture });
            int count = 0;
            long sum = 0;
            await foreach (TypedRecord record in reader.GetRecordsAsync<TypedRecord>()) {
                if (_validate) TypedWorkbookFixture.ValidateRecord(record, count + 1);
                sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record)); count++;
            }
            return CsvBenchmarkFixture.Check(sum, count, RowCount, _expected);
        }
    }
}
