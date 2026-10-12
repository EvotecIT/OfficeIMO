using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using OfficeIMO.CSV;
using OfficeIMO.Data;
using System.Buffers.Text;
using System.Data.Common;
using System.Globalization;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;
using SylvanCsv = Sylvan.Data.Csv.CsvDataReader;
using SylvanOptions = Sylvan.Data.Csv.CsvDataReaderOptions;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>The upstream headerless four-field CSV workload with a borrowed name field.</summary>
    internal class CsvReadWorkload {
        protected byte[] Data = [];
        private long _expected;
        public int RowCount { get; set; } = 50_000;
        public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();

        public void Setup() {
            Data = CsvBenchmarkFixture.Create(RowCount, headers: false);
            _expected = TypedWorkbookFixture.ExpectedChecksum(RowCount);
            CsvBenchmarkFixture.Describe(Data, RowCount, "raw original headerless");
            using (MemoryStream stream = new MemoryStream(Data, false))
            using (DbDataReader reader = CsvDocument.OpenDataReader(stream, new CsvLoadOptions { HasHeaderRow = false }))
                CsvBenchmarkFixture.Validate(reader, RowCount, headers: false);
            using (MemoryStream stream = new MemoryStream(Data, false))
            using (StreamReader text = new StreamReader(stream))
            using (SylvanCsv reader = SylvanCsv.Create(text, new SylvanOptions { HasHeaders = false, Culture = CultureInfo.InvariantCulture }))
                CsvBenchmarkFixture.Validate(reader, RowCount, headers: false);
            ValidateExcelReader();
            ExcelReaderOriginal();
#if OFFICEIMO_BENCHMARK_NEW_APIS
            OfficeIMOBorrowedUtf8();
#endif
            SylvanOriginal();
        }

        public long ExcelReaderOriginal() => ReadExcelReader(materialize: false);

#if OFFICEIMO_BENCHMARK_NEW_APIS
        public long OfficeIMOBorrowedUtf8() {
            using MemoryStream stream = new MemoryStream(Data, false);
            using DbDataReader reader = CsvDocument.OpenDataReader(stream, new CsvLoadOptions { HasHeaderRow = false });
            int count = 0;
            long sum = 0;
            while (reader.Read()) {
                if (!reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> name) || !reader.TryGetUtf8Text(1, out ReadOnlySpan<byte> idText)
                    || !reader.TryGetUtf8Text(2, out ReadOnlySpan<byte> dateText) || !reader.TryGetUtf8Text(3, out ReadOnlySpan<byte> valueText)
                    || !int.TryParse(idText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id)
                    || !Utf8Parser.TryParse(dateText, out DateTime date, out int consumed, 'O') || consumed != dateText.Length
                    || !double.TryParse(valueText, NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
                    throw new InvalidDataException("Ordinary CSV fields must provide their normalized UTF-8 text and parse exactly.");
                sum = unchecked(sum + name.Length + id + date.Ticks + (long)value);
                count++;
            }
            return CsvBenchmarkFixture.Check(sum, count, RowCount, _expected);
        }
#endif

        public long SylvanOriginal() => ReadSylvan(materialize: false);

        internal long ReadOfficeIMOMaterialized() {
            using MemoryStream stream = new MemoryStream(Data, false);
            using DbDataReader reader = CsvDocument.OpenDataReader(stream, new CsvLoadOptions { HasHeaderRow = false });
            int count = 0;
            long sum = 0;
            while (reader.Read()) {
                sum = unchecked(sum + reader.GetString(0).Length + reader.GetInt32(1)
                    + reader.GetDateTime(2).Ticks + (long)reader.GetDouble(3));
                count++;
            }
            return CsvBenchmarkFixture.Check(sum, count, RowCount, _expected);
        }

        internal long ReadExcelReader(bool materialize) {
            using MemoryStream stream = new MemoryStream(Data, false);
            using CsvReader workbook = ExcelReaderApi.FromCsv(stream);
            int count = 0;
            long sum = 0;
            foreach (Row row in workbook.FirstSheet) {
                if (!row[1].TryParse(CultureInfo.InvariantCulture, out int id)
                    || !row[2].TryGetDateTime(out DateTime date)
                    || !row[3].TryParse(CultureInfo.InvariantCulture, out double value))
                    throw new InvalidDataException("CSV scalar conversion failed.");
                sum = unchecked(sum + (materialize ? row[0].GetString().Length : row[0].Value.Length) + id + date.Ticks + (long)value);
                count++;
            }
            return CsvBenchmarkFixture.Check(sum, count, RowCount, _expected);
        }

        internal long ReadSylvan(bool materialize) {
            using MemoryStream stream = new MemoryStream(Data, false);
            using StreamReader text = new StreamReader(stream);
            using SylvanCsv reader = SylvanCsv.Create(text, new SylvanOptions { HasHeaders = false, Culture = CultureInfo.InvariantCulture });
            int count = 0;
            long sum = 0;
            while (reader.Read()) {
                sum = unchecked(sum + (materialize ? reader.GetString(0).Length : reader.GetFieldSpan(0).Length)
                    + reader.GetInt32(1) + reader.GetDateTime(2).Ticks + (long)reader.GetDouble(3));
                count++;
            }
            return CsvBenchmarkFixture.Check(sum, count, RowCount, _expected);
        }

        internal void SetupMaterialized() {
            Data = CsvBenchmarkFixture.Create(RowCount, headers: false);
            _expected = TypedWorkbookFixture.ExpectedChecksum(RowCount);
            CsvBenchmarkFixture.Describe(Data, RowCount, "raw materialized name");
            ValidateExcelReader();
            using (MemoryStream stream = new MemoryStream(Data, false))
            using (DbDataReader reader = CsvDocument.OpenDataReader(stream, new CsvLoadOptions { HasHeaderRow = false }))
                CsvBenchmarkFixture.Validate(reader, RowCount, headers: false);
            using (MemoryStream stream = new MemoryStream(Data, false))
            using (StreamReader text = new StreamReader(stream))
            using (SylvanCsv reader = SylvanCsv.Create(text, new SylvanOptions { HasHeaders = false, Culture = CultureInfo.InvariantCulture }))
                CsvBenchmarkFixture.Validate(reader, RowCount, headers: false);
        }

        private void ValidateExcelReader() {
            using MemoryStream stream = new MemoryStream(Data, false);
            using CsvReader workbook = ExcelReaderApi.FromCsv(stream);
            int count = 0;
            foreach (Row row in workbook.FirstSheet) {
                if (row.ColumnCount != 4 || !row[1].TryParse(CultureInfo.InvariantCulture, out int id)
                    || !row[2].TryGetDateTime(out DateTime date)
                    || !row[3].TryParse(CultureInfo.InvariantCulture, out double value))
                    throw new InvalidDataException("CSV field shape or conversion differs.");
                TypedWorkbookFixture.ValidateRecord(new TypedRecord { Name = row[0].GetString(), Id = id, Date = date, Value = value }, ++count);
            }
            if (count != RowCount) throw new InvalidDataException("CSV row count differs.");
        }
    }

#if OFFICEIMO_BENCHMARK_NEW_APIS
    [MemoryDiagnoser]
    [BenchmarkCategory("CsvRawOriginal")]
    public class CsvRawReadBenchmarks {
        private readonly CsvReadWorkload _workload = new();
        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 50_000;
        public IEnumerable<int> RowCounts() => _workload.RowCounts();
        [GlobalSetup]
        public void Setup() { _workload.RowCount = RowCount; _workload.Setup(); }
        [Benchmark(Baseline = true)]
        public long ExcelReaderOriginal() => _workload.ExcelReaderOriginal();
        [Benchmark]
        public long OfficeIMOBorrowedUtf8() => _workload.OfficeIMOBorrowedUtf8();
        [Benchmark]
        public long SylvanOriginal() => _workload.SylvanOriginal();
    }
#endif

    /// <summary>Every engine materializes the name while parsing the remaining scalar fields.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("CsvRawMaterialized")]
    public class CsvMaterializedReadBenchmarks {
        private readonly CsvReadWorkload _workload = new();
        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 50_000;
        public IEnumerable<int> RowCounts() => _workload.RowCounts();
        [GlobalSetup]
        public void Setup() {
            _workload.RowCount = RowCount;
            _workload.SetupMaterialized();
            ExcelReaderMaterialized(); OfficeIMOMaterialized(); SylvanMaterialized();
        }
        [Benchmark(Baseline = true)]
        public long ExcelReaderMaterialized() => _workload.ReadExcelReader(materialize: true);
        [Benchmark]
        public long OfficeIMOMaterialized() => _workload.ReadOfficeIMOMaterialized();
        [Benchmark]
        public long SylvanMaterialized() => _workload.ReadSylvan(materialize: true);
    }
}
