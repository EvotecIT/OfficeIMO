using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer.Xlsx;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Writes the prepared upstream records through each library's ordinary XLSX API.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("WriteOriginal")]
    public class WriterBenchmarks {
        private List<TypedRecord> _records = [];

        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 50_000;

        public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();

        [GlobalSetup]
        public async Task SetupAsync() {
            BenchmarkInput.WriteDescription();
            _records = new List<TypedRecord>(RowCount);
            for (int index = 1; index <= RowCount; index++) _records.Add(TypedWorkbookFixture.ExpectedRecord(index));
            using MemoryStream officeStream = new MemoryStream(4 * 1024 * 1024);
            WriteOfficeIMO(officeStream);
            await using MemoryStream peerStream = new MemoryStream(4 * 1024 * 1024);
            await WriteExcelReaderAsync(peerStream);
            WrittenWorkbookValidation.Validate(officeStream.ToArray(), RowCount, officeIMO: true);
            WrittenWorkbookValidation.Validate(peerStream.ToArray(), RowCount, officeIMO: false);
        }

        [Benchmark]
        public long OfficeIMO() {
            using MemoryStream stream = new MemoryStream(4 * 1024 * 1024);
            var destination1 = stream;

            var writeOptions1 = new ExcelWriterOptions {
                UseSharedStrings = false,
                CreateTable = false
            };

            using (var workbookWriter1 = ExcelDocument.CreateWriter(destination1, writeOptions1, leaveOpen: true)) {
                workbookWriter1.WriteTable(writeOptions1.SheetName, _records, ["Name", "Id", "Date", "Value"], static (row, record) => {
                    row.Write(record.Name);
                    row.Write(record.Id);
                    row.Write(record.Date);
                    row.Write(record.Value);
                });
            }

            if (destination1.CanSeek)
                destination1.Position = 0;
            return stream.Length;
        }

        [Benchmark(Baseline = true)]
        public async Task<long> ExcelReaderWriter() {
            await using MemoryStream stream = new MemoryStream(4 * 1024 * 1024);
            await using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(stream, leaveOpen: true)) {
                XlsxSheetWriter sheet = workbook.AddSheet("S1");
                using (XlsxRowWriter header = sheet.StartRow()) {
                    header.Write("Name");
                    header.Write("Id");
                    header.Write("Date");
                    header.Write("Value");
                }
                for (int index = 0; index < _records.Count; index++) {
                    TypedRecord record = _records[index];
                    using XlsxRowWriter row = sheet.StartRow();
                    row.Write(record.Name);
                    row.Write(record.Id);
                    row.Write(record.Date);
                    row.Write(record.Value);
                }
                await sheet.EndAsync();
                await workbook.EndAsync();
            }
            return stream.Length;
        }

        private void WriteOfficeIMO(Stream stream) {
            var destination2 = stream;

            var writeOptions2 = new ExcelWriterOptions {
                UseSharedStrings = false,
                CreateTable = false
            };

            using (var workbookWriter2 = ExcelDocument.CreateWriter(destination2, writeOptions2, leaveOpen: true)) {
                workbookWriter2.WriteTable(writeOptions2.SheetName, _records, ["Name", "Id", "Date", "Value"], static (row, record) => {
                    row.Write(record.Name);
                    row.Write(record.Id);
                    row.Write(record.Date);
                    row.Write(record.Value);
                });
            }

            if (destination2.CanSeek)
                destination2.Position = 0;
        }

        private async Task WriteExcelReaderAsync(Stream stream) {
            await using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(stream, leaveOpen: true);
            XlsxSheetWriter sheet = workbook.AddSheet("S1");
            using (XlsxRowWriter header = sheet.StartRow()) {
                header.Write("Name");
                header.Write("Id");
                header.Write("Date");
                header.Write("Value");
            }
            for (int index = 0; index < _records.Count; index++) {
                TypedRecord record = _records[index];
                using XlsxRowWriter row = sheet.StartRow();
                row.Write(record.Name);
                row.Write(record.Id);
                row.Write(record.Date);
                row.Write(record.Value);
            }
            await sheet.EndAsync();
            await workbook.EndAsync();
        }
    }
}
