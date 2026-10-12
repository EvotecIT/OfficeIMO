using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using OfficeIMO.CSV;
using OfficeIMO.Data;
using System.Data.Common;
using System.Globalization;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;
using SylvanCsv = Sylvan.Data.Csv.CsvDataReader;
using SylvanOptions = Sylvan.Data.Csv.CsvDataReaderOptions;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>32 headerless ASCII text fields, with separate borrowed and materialized work.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("CsvWide")]
    public class CsvWideReadBenchmarks {
        private byte[] _data = [];
        private long _expected;
        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 50_000;
        [ParamsSource(nameof(MaterializationModes))]
        public bool MaterializeStrings { get; set; }
        public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();
        public IEnumerable<bool> MaterializationModes() {
#if OFFICEIMO_BENCHMARK_NEW_APIS
            yield return false;
#endif
            yield return true;
        }

        [GlobalSetup]
        public void Setup() {
            _data = CsvBenchmarkFixture.CreateWide(RowCount);
            _expected = 0;
            for (int row = 0; row < RowCount; row++)
                for (int column = 0; column < 32; column++) _expected += CsvBenchmarkFixture.Names[(row + column) % 8].Length;
            CsvBenchmarkFixture.Describe(_data, RowCount, $"32-column wide, materializeStrings={MaterializeStrings}");
            using (MemoryStream stream = new MemoryStream(_data, false))
            using (DbDataReader reader = CsvDocument.OpenDataReader(stream, new CsvLoadOptions { HasHeaderRow = false })) {
                int row = 0;
                while (reader.Read()) {
                    if (reader.FieldCount != 32) throw new InvalidDataException("Wide CSV width differs.");
                    for (int column = 0; column < 32; column++) {
                        if (reader.GetString(column) != CsvBenchmarkFixture.Names[(row + column) % 8])
                            throw new InvalidDataException("Wide CSV value differs.");
                    }
                    row++;
                }
                if (row != RowCount) throw new InvalidDataException("Wide CSV row count differs.");
            }
            ValidatePeerFields();
            ExcelReaderWide(); OfficeIMOWide(); SylvanWide();
        }

        [Benchmark(Baseline = true)]
        public long ExcelReaderWide() {
            using MemoryStream stream = new MemoryStream(_data, false);
            using CsvReader workbook = ExcelReaderApi.FromCsv(stream);
            long sum = 0;
            int count = 0;
            foreach (Row row in workbook.FirstSheet) {
                for (int column = 0; column < row.ColumnCount; column++)
                    sum += MaterializeStrings ? row[column].GetString().Length : row[column].Value.Length;
                count++;
            }
            return Check(sum, count);
        }

        [Benchmark]
        public long OfficeIMOWide() {
            using MemoryStream stream = new MemoryStream(_data, false);
            using DbDataReader reader = CsvDocument.OpenDataReader(stream, new CsvLoadOptions { HasHeaderRow = false });
            long sum = 0;
            int count = 0;
            while (reader.Read()) {
                for (int column = 0; column < reader.FieldCount; column++) {
#if OFFICEIMO_BENCHMARK_NEW_APIS
                    if (!MaterializeStrings) {
                        if (!reader.TryGetUtf8Text(column, out ReadOnlySpan<byte> text)) throw new InvalidDataException("Expected borrowed CSV text.");
                        sum += text.Length;
                    } else
#endif
                    sum += reader.GetString(column).Length;
                }
                count++;
            }
            return Check(sum, count);
        }

        [Benchmark]
        public long SylvanWide() {
            using MemoryStream stream = new MemoryStream(_data, false);
            using StreamReader text = new StreamReader(stream);
            using SylvanCsv reader = SylvanCsv.Create(text, new SylvanOptions { HasHeaders = false, Culture = CultureInfo.InvariantCulture });
            long sum = 0;
            int count = 0;
            while (reader.Read()) {
                for (int column = 0; column < reader.FieldCount; column++)
                    sum += MaterializeStrings ? reader.GetString(column).Length : reader.GetFieldSpan(column).Length;
                count++;
            }
            return Check(sum, count);
        }

        private long Check(long checksum, int count) => checksum == _expected && count == RowCount
            ? checksum : throw new InvalidDataException("Wide CSV count or checksum differs.");

        private void ValidatePeerFields() {
            using (MemoryStream stream = new MemoryStream(_data, false))
            using (CsvReader workbook = ExcelReaderApi.FromCsv(stream)) {
                int index = 0;
                foreach (Row row in workbook.FirstSheet) {
                    if (row.ColumnCount != 32) throw new InvalidDataException("Peer wide width differs.");
                    for (int column = 0; column < 32; column++) {
                        if (row[column].GetString() != CsvBenchmarkFixture.Names[(index + column) % 8])
                            throw new InvalidDataException("Peer wide value differs.");
                    }
                    index++;
                }
                if (index != RowCount) throw new InvalidDataException("Peer wide row count differs.");
            }
            using (MemoryStream stream = new MemoryStream(_data, false))
            using (StreamReader text = new StreamReader(stream))
            using (SylvanCsv reader = SylvanCsv.Create(text, new SylvanOptions { HasHeaders = false, Culture = CultureInfo.InvariantCulture })) {
                int index = 0;
                while (reader.Read()) {
                    if (reader.FieldCount != 32) throw new InvalidDataException("Sylvan wide width differs.");
                    for (int column = 0; column < 32; column++) {
                        if (reader.GetString(column) != CsvBenchmarkFixture.Names[(index + column) % 8])
                            throw new InvalidDataException("Sylvan wide value differs.");
                    }
                    index++;
                }
                if (index != RowCount) throw new InvalidDataException("Sylvan wide row count differs.");
            }
        }
    }
}
