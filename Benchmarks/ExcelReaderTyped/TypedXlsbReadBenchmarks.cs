using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader.Xlsb;
using Sylvan.Data;
using Sylvan.Data.Excel;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Materializes typed records from the upstream header-bearing XLSB fixture.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("TypedXlsb")]
    public class TypedXlsbReadBenchmarks {
        internal byte[] Workbook { get; private set; } = [];
        private long _expected;

        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 50_000;

        public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();

        [GlobalSetup]
        public async Task SetupAsync() {
            BenchmarkInput.WriteDescription();
            Workbook = await TypedWorkbookFixture.CreateXlsbAsync(RowCount);
            _expected = TypedWorkbookFixture.ExpectedChecksum(RowCount);
            using (MemoryStream stream = new MemoryStream(Workbook, writable: false)) {
                using XlsbWorkbook workbook = ExcelReaderApi.FromXlsb(stream);
                if (workbook.SheetCount != 1) throw new InvalidDataException("Incorrect typed XLSB sheet count.");
                Validate(ExcelParser.FromAttributes<TypedRecord>().Parse(workbook.FirstSheet));
            }
            using (ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(Workbook, new ExcelReadOptions { HasHeaderRow = true })) {
                ValidateHeaders(reader);
                int count = 0;
                while (reader.Read()) TypedWorkbookFixture.ValidateRecord(new TypedRecord {
                    Name = reader.GetString(0), Id = reader.GetInt32(1), Date = reader.GetDateTime(2), Value = reader.GetDouble(3),
                }, ++count);
                if (count != RowCount || reader.NextResult()) throw new InvalidDataException("Incorrect typed XLSB row/sheet count.");
            }
            using (MemoryStream stream = new MemoryStream(Workbook, writable: false)) {
                using ExcelDataReader reader = global::Sylvan.Data.Excel.ExcelDataReader.Create(stream, ExcelWorkbookType.ExcelBinary, new ExcelDataReaderOptions());
                ValidateHeaders(reader);
                Validate(reader.GetRecords<TypedRecord>());
                if (reader.NextResult()) throw new InvalidDataException("Incorrect typed XLSB sheet count.");
            }
            ExcelReaderTypedXlsb();
            OfficeIMOTypedXlsb();
            SylvanTypedXlsb();
            Console.WriteLine($"Validated typed XLSB: rows={RowCount}, packageBytes={Workbook.Length}, checksum={_expected}.");
        }

        [Benchmark(Baseline = true)]
        public long ExcelReaderTypedXlsb() {
            using MemoryStream stream = new MemoryStream(Workbook, writable: false);
            using XlsbWorkbook workbook = ExcelReaderApi.FromXlsb(stream);
            long sum = 0;
            int count = 0;
            foreach (TypedRecord record in ExcelParser.FromAttributes<TypedRecord>().Parse(workbook.FirstSheet)) {
                sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record));
                count++;
            }
            return Check(sum, count);
        }

        [Benchmark]
        public long OfficeIMOTypedXlsb() {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(Workbook, new ExcelReadOptions { HasHeaderRow = true });
            long sum = 0;
            int count = 0;
            while (reader.Read()) {
                sum = unchecked(sum + TypedWorkbookFixture.Accumulate(new TypedRecord {
                    Name = reader.GetString(0), Id = reader.GetInt32(1), Date = reader.GetDateTime(2), Value = reader.GetDouble(3),
                }));
                count++;
            }
            return Check(sum, count);
        }

        [Benchmark]
        public long SylvanTypedXlsb() {
            using MemoryStream stream = new MemoryStream(Workbook, writable: false);
            using ExcelDataReader reader = global::Sylvan.Data.Excel.ExcelDataReader.Create(stream, ExcelWorkbookType.ExcelBinary, new ExcelDataReaderOptions());
            long sum = 0;
            int count = 0;
            foreach (TypedRecord record in reader.GetRecords<TypedRecord>()) {
                sum = unchecked(sum + TypedWorkbookFixture.Accumulate(record));
                count++;
            }
            return Check(sum, count);
        }

        internal void Validate(IEnumerable<TypedRecord> records) {
            int count = 0;
            foreach (TypedRecord record in records) TypedWorkbookFixture.ValidateRecord(record, ++count);
            if (count != RowCount) throw new InvalidDataException("Incorrect typed XLSB row count.");
        }

        internal long Check(long sum, int count) => sum == _expected && count == RowCount
            ? sum : throw new InvalidDataException("Incorrect typed XLSB count or checksum.");

        private static void ValidateHeaders(System.Data.Common.DbDataReader reader) {
            if (reader.FieldCount != 4) throw new InvalidDataException("Incorrect typed XLSB width.");
            for (int column = 0; column < 4; column++) {
                if (reader.GetName(column) != TypedWorkbookFixture.Headers[column]) throw new InvalidDataException("Incorrect typed XLSB header.");
            }
        }
    }
}
