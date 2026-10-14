using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsx;
using OfficeIMO.Data;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>A distinct model whose mapping caches are untouched by fixture creation.</summary>
    public sealed class ColdStartRecord {
        public string? Name { get; set; }
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public double Value { get; set; }
    }

    /// <summary>Measures one first typed mapping use in each of sixteen fresh processes.</summary>
    [MemoryDiagnoser]
    [SimpleJob(RunStrategy.ColdStart, launchCount: 16, warmupCount: 0, iterationCount: 1, invocationCount: 1)]
    [BenchmarkCategory("ColdStart")]
    public class ColdStartReadBenchmarks {
        private byte[] _workbook = [];

        [GlobalSetup]
        public async Task SetupAsync() {
            BenchmarkInput.WriteDescription();
            _workbook = await ColdStartWorkload.BuildInputAsync();
            // Reader methods and reflection mapping must not run during cold setup.
            TypedWorkbookFixture.ValidateShape(_workbook, ColdStartWorkload.Rows, "Original");
        }

        [Benchmark(Baseline = true)]
        public long ExcelReaderAttributes() {
            using XlsxWorkbook workbook = ExcelReaderApi.FromXlsx(new MemoryStream(_workbook, writable: false));
            long sum = 0;
            foreach (ColdStartRecord record in ExcelParser.FromAttributes<ColdStartRecord>().Parse(workbook.FirstSheet)) sum += record.Id;
            return ColdStartWorkload.Check(sum);
        }

        [Benchmark]
        public long OfficeIMOAutomaticMapping() {
            using MemoryStream stream = new MemoryStream(_workbook, writable: false);
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(stream);
            long sum = 0;
            foreach (ColdStartRecord record in reader.RowsAs<ColdStartRecord>()) sum += record.Id;
            return ColdStartWorkload.Check(sum);
        }

        [Benchmark]
        public long ExcelReaderFluentMapping() {
            using XlsxWorkbook workbook = ExcelReaderApi.FromXlsx(new MemoryStream(_workbook, writable: false));
            ExcelParser<ColdStartRecord> parser = CreateFluentParser();
            long sum = 0;
            foreach (ColdStartRecord record in parser.Parse(workbook.FirstSheet)) sum += record.Id;
            return ColdStartWorkload.Check(sum);
        }

        [Benchmark]
        public long ExcelReaderAttributeFallback() {
            using XlsxWorkbook workbook = ExcelReaderApi.FromXlsx(new MemoryStream(_workbook, writable: false));
            ExcelParser<ColdStartRecord> parser = CreateAttributeFallbackParser();
            long sum = 0;
            foreach (ColdStartRecord record in parser.Parse(workbook.FirstSheet)) sum += record.Id;
            return ColdStartWorkload.Check(sum);
        }

        internal void Validate() {
            foreach (ExcelParser<ColdStartRecord> parser in new[] {
                ExcelParser.FromAttributes<ColdStartRecord>(), CreateFluentParser(), CreateAttributeFallbackParser()
            }) {
                using XlsxWorkbook workbook = ExcelReaderApi.FromXlsx(new MemoryStream(_workbook, writable: false));
                int rows = 0;
                foreach (ColdStartRecord record in parser.Parse(workbook.FirstSheet)) ColdStartWorkload.ValidateRecord(record, ++rows);
                if (rows != ColdStartWorkload.Rows) throw new InvalidDataException("Cold parser strategy count differs.");
            }
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(_workbook);
            int count = 0;
            foreach (ColdStartRecord record in reader.RowsAs<ColdStartRecord>()) ColdStartWorkload.ValidateRecord(record, ++count);
            if (count != ColdStartWorkload.Rows) throw new InvalidDataException("Cold projection count differs.");
            ExcelReaderAttributes(); OfficeIMOAutomaticMapping(); ExcelReaderFluentMapping(); ExcelReaderAttributeFallback();
        }

        private static ExcelParser<ColdStartRecord> CreateFluentParser() =>
            ExcelParser.Build<ColdStartRecord>(static builder => builder
                .Factory(static () => new ColdStartRecord())
                .Property(["Name"], ExcelCellReaders.String, static (ref record, value) => record.Name = value)
                .Property(["Id"], ExcelCellReaders.Parsable, static (ref ColdStartRecord record, int value) => record.Id = value)
                .Property(["Date"], ExcelCellReaders.DateTimeSerial, static (ref record, value) => record.Date = value)
                .Property(["Value"], ExcelCellReaders.Parsable, static (ref ColdStartRecord record, double value) => record.Value = value));

        private static ExcelParser<ColdStartRecord> CreateAttributeFallbackParser() =>
            ExcelParser.BuildWithAttributeFallback<ColdStartRecord>(static builder => builder
                .Property(["Id"], ExcelCellReaders.Parsable, static (ref ColdStartRecord record, int value) => record.Id = value));
    }

    internal static class ColdStartWorkload {
        internal const int Rows = 200;
        internal static long Check(long sum) => sum == (long)Rows * (Rows + 1) / 2 ? sum : throw new InvalidDataException("Cold parse checksum differs.");

        internal static void ValidateRecord(ColdStartRecord record, int row) => TypedWorkbookFixture.ValidateRecord(new TypedRecord {
            Name = record.Name, Id = record.Id, Date = record.Date, Value = record.Value
        }, row);

        internal static List<ColdStartRecord> Records() => Enumerable.Range(1, Rows).Select(index => {
            TypedRecord record = TypedWorkbookFixture.ExpectedRecord(index);
            return new ColdStartRecord { Name = record.Name, Id = record.Id, Date = record.Date, Value = record.Value };
        }).ToList();

        internal static async Task<byte[]> BuildInputAsync() {
            await using MemoryStream stream = new MemoryStream();
            await using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(stream, leaveOpen: true)) {
                await using XlsxSheetWriter sheet = workbook.AddSheet("S1");
                await using (XlsxRowWriter header = await sheet.StartRowAsync()) {
                    foreach (string name in TypedWorkbookFixture.Headers) header.Write(name);
                }
                foreach (ColdStartRecord record in Records()) {
                    await using XlsxRowWriter row = await sheet.StartRowAsync();
                    row.Write(record.Name); row.Write(record.Id); row.Write(record.Date); row.Write(record.Value);
                }
                await sheet.EndAsync();
                await workbook.EndAsync();
            }
            byte[] bytes = stream.ToArray();
            BenchmarkInput.WriteFixtureIdentity($"read/cold/Xlsx/dataRows={Rows}", bytes);
            return bytes;
        }
    }
}
