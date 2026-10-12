using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer.Csv;
using OfficeIMO.CSV;
using Sylvan.Data;
using Sylvan.Data.Csv;
using System.Data.Common;
using System.Globalization;
using System.Text;
using PeerCsvWriter = ExcelReader.Core.Writer.Csv.CsvWorkbookWriter;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Serialize the same materialized records and qualify all fields through independent readers.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("CsvWrite")]
    public class CsvWriterBenchmarks {
        private List<TypedRecord> _records = [];
        private readonly long[] _lengths = new long[3];
        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 50_000;
        public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();

        [GlobalSetup]
        public void Setup() {
            _records = Enumerable.Range(1, RowCount).Select(TypedWorkbookFixture.ExpectedRecord).ToList();
            for (int engine = 0; engine < 3; engine++) {
                using MemoryStream stream = new MemoryStream(4 * 1024 * 1024);
                Write(stream, engine);
                _lengths[engine] = stream.Length;
                stream.Position = 0;
                using (StreamReader text = new StreamReader(stream, Encoding.UTF8, true, 1024, leaveOpen: true))
                using (CsvDataReader reader = global::Sylvan.Data.Csv.CsvDataReader.Create(text,
                    new global::Sylvan.Data.Csv.CsvDataReaderOptions { Culture = CultureInfo.InvariantCulture }))
                    CsvBenchmarkFixture.Validate(reader, RowCount, headers: true);
                stream.Position = 0;
                using (DbDataReader reader = CsvDocument.OpenDataReader(stream)) CsvBenchmarkFixture.Validate(reader, RowCount, headers: true);
                Console.WriteLine($"CSV writer engine={engine}, rows={RowCount}, bytes={stream.Length}, all fields validated.");
            }
            CsvBenchmarkFixture.Describe(CsvBenchmarkFixture.Create(RowCount, true), RowCount, "writer records");
            ExcelReaderWriter(); OfficeIMOWriteObjects(); SylvanWriter();
        }

        [Benchmark(Baseline = true)]
        public long ExcelReaderWriter() => MeasureWrite(0);
        [Benchmark]
        public long OfficeIMOWriteObjects() => MeasureWrite(1);
        [Benchmark]
        public long SylvanWriter() => MeasureWrite(2);

        private long MeasureWrite(int engine) {
            using MemoryStream stream = new MemoryStream(4 * 1024 * 1024);
            Write(stream, engine);
            return stream.Length == _lengths[engine] ? stream.Length : throw new InvalidDataException("CSV output length differs after qualification.");
        }

        private void Write(Stream stream, int engine) {
            if (engine == 0) {
                using PeerCsvWriter workbook = PeerCsvWriter.Create(stream, leaveOpen: true);
                using CsvSheetWriter sheet = workbook.AddSheet("S1");
                using (ExcelReader.Core.Writer.Csv.CsvRowWriter header = sheet.StartRow()) {
                    foreach (string column in TypedWorkbookFixture.Headers) header.Write(column);
                }
                foreach (TypedRecord record in _records) {
                    using ExcelReader.Core.Writer.Csv.CsvRowWriter row = sheet.StartRow();
                    row.Write(record.Name); row.Write(record.Id); row.Write(record.Date); row.Write(record.Value);
                }
            } else if (engine == 1) {
                using StreamWriter text = new StreamWriter(stream, new UTF8Encoding(false), 1024, leaveOpen: true);
                CsvDocument.WriteObjects(text, _records, new CsvSaveOptions {
                    Culture = CultureInfo.InvariantCulture, DateTimeFormat = "O",
                });
            } else {
                using StreamWriter text = new StreamWriter(stream, new UTF8Encoding(false), 1024, leaveOpen: true);
                using CsvDataWriter writer = global::Sylvan.Data.Csv.CsvDataWriter.Create(text);
                using DbDataReader reader = ObjectDataReader.Create(_records);
                writer.Write(reader);
            }
        }
    }
}
