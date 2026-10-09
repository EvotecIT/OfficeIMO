#if OFFICEIMO_BENCHMARK_ARROW && OFFICEIMO_BENCHMARK_NEW_APIS
using Apache.Arrow;
using Apache.Arrow.Types;
using BenchmarkDotNet.Attributes;
using ExcelReader.Arrow;
using ExcelReader.Core.Writer.Xlsx;
using System.IO.Compression;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Writes the same prepared string batch through public workbook APIs and caller orchestration.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("WriteArrowStrings")]
    public class ArrowStringWriterBenchmarks {
        private const int Rows = 100_000, Columns = 4;
        private readonly MemoryStream _output = new(32 * 1024 * 1024);
        private RecordBatch? _batch;
        private string[] _headers = [];

        [GlobalSetup]
        public void Setup() {
            BenchmarkInput.WriteDescription();
            try {
                _headers = Enumerable.Range(0, Columns).Select(column => "c" + column).ToArray();
                List<IArrowArray> arrays = new List<IArrowArray>(Columns);
                try {
                    Schema.Builder schema = new Schema.Builder();
                    for (int column = 0; column < Columns; column++) {
                        StringArray.Builder values = new StringArray.Builder();
                        for (int index = 0; index < Rows; index++) values.Append(ExpectedValue(column, index));
                        arrays.Add(values.Build());
                        schema.Field(new Field(_headers[column], StringType.Default, nullable: false));
                    }
                    _batch = new RecordBatch(schema.Build(), arrays, Rows);
                } catch {
                    foreach (IArrowArray array in arrays) array.Dispose();
                    throw;
                }
                if (_batch.Schema.FieldsList.Count != Columns || _batch.Length != Rows
                    || _batch.Schema.FieldsList.Any(field => field.IsNullable || field.DataType.TypeId != ArrowTypeId.String)
                    || Enumerable.Range(0, Columns).Any(column => _batch.Column(column).NullCount != 0))
                    throw new InvalidDataException("The prepared Arrow string batch differs from the upstream schema.");
                for (int column = 0; column < Columns; column++) {
                    IReadOnlyList<ArrowBuffer> buffers = _batch.Column(column).Data.Buffers;
                    for (int buffer = 0; buffer < buffers.Count; buffer++)
                        BenchmarkInput.WriteFixtureIdentity($"write/ArrowStrings/rows={Rows}/column={column}/buffer={buffer}",
                            buffers[buffer].Memory.Span);
                }
                OfficeIMOPublicRowOrchestration();
                Validate(_output.ToArray(), "OfficeIMO public row orchestration");
                ExcelReaderRecordBatch();
                Validate(_output.ToArray(), "ExcelReader public WriteRecordBatch");
            } catch {
                Cleanup();
                throw;
            }
        }

        [GlobalCleanup]
        public void Cleanup() {
            _batch?.Dispose();
            _batch = null;
            _output.Dispose();
        }

        [Benchmark(Baseline = true)]
        public long ExcelReaderRecordBatch() {
            _output.SetLength(0);
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(_output, leaveOpen: true))
                workbook.WriteRecordBatch(_batch!);
            return _output.Length;
        }

        [Benchmark]
        public long OfficeIMOPublicRowOrchestration() {
            _output.SetLength(0);
            ExcelDocument.WriteRows(_output, Enumerable.Range(0, Rows), _headers, WriteOfficeIMORow,
                new ExcelTabularWriteOptions { SheetName = "Sheet1", UseSharedStrings = false });
            return _output.Length;
        }

        private void WriteOfficeIMORow(ExcelDocument.ExcelTabularRowWriter row, int index) {
            for (int column = 0; column < Columns; column++) {
                ReadOnlySpan<byte> text = ((StringArray)_batch!.Column(column)).GetBytes(index, out bool isNull);
                if (isNull) throw new InvalidDataException("The required string column contains null.");
                row.WriteUtf8(text);
            }
        }

        private static string ExpectedValue(int column, int index) => $"customer-{column}-{index % 5000:D5}";

        private void Validate(byte[] bytes, string engine) {
            using (MemoryStream stream = new MemoryStream(bytes, writable: false)) {
                using ZipArchive package = new ZipArchive(stream, ZipArchiveMode.Read);
                WrittenWorkbookValidation.ValidateEntries(package);
                WrittenWorkbookValidation.ValidateRelationships(package);
            }
            WorkbookScanWorkload scan = new WorkbookScanWorkload();
            scan.Setup(bytes, ComparisonWorkbookFormat.Xlsx, Rows + 1, Columns, index => index == 0
                ? _headers.Cast<object>().ToArray()
                : Enumerable.Range(0, Columns).Select(column => (object)ExpectedValue(column, index - 1)).ToArray());
            Console.WriteLine($"Validated {engine}: Arrow rows={Rows}, required string columns={Columns}, "
                + $"outputRows={Rows + 1}, header=true, bytes={bytes.Length}, retainedOutputCapacity=32MiB; "
                + "OfficeIMO includes ordinary Enumerable.Range/callback/column access costs.");
        }
    }
}
#endif