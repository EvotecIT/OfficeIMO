#if OFFICEIMO_BENCHMARK_NEW_APIS
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer.Xlsx;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Compares public UTF-8 row APIs on the upstream unmanaged string-column input.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("WriteUtf8Public")]
    public unsafe class Utf8WriterBenchmarks {
        private const int Rows = 100_000, Columns = 4;
        private readonly MemoryStream _output = new(32 * 1024 * 1024);
        private readonly List<IntPtr> _native = [];
        private Utf8Column[] _columns = [];

        [GlobalSetup]
        public void Setup() {
            BenchmarkInput.WriteDescription();
            try {
                _columns = new Utf8Column[Columns];
                for (int column = 0; column < Columns; column++) {
                    int[] offsets = new int[Rows + 1];
                    using MemoryStream data = new MemoryStream();
                    for (int row = 0; row < Rows; row++) {
                        byte[] utf8 = Encoding.UTF8.GetBytes(ExpectedValue(column, row));
                        data.Write(utf8);
                        offsets[row + 1] = checked((int)data.Length);
                    }
                    byte[] values = data.ToArray();
                    BenchmarkInput.WriteFixtureIdentity($"write/Utf8/Xlsx/rows={Rows}/column={column}/offsets",
                        MemoryMarshal.AsBytes(offsets.AsSpan()));
                    BenchmarkInput.WriteFixtureIdentity($"write/Utf8/Xlsx/rows={Rows}/column={column}/values", values);
                    _columns[column] = new Utf8Column(Pin(MemoryMarshal.AsBytes(offsets.AsSpan())), Pin(values));
                }
                OfficeIMOUtf8();
                Validate(_output.ToArray(), "OfficeIMO");
                ExcelReaderPublicUtf8();
                Validate(_output.ToArray(), "ExcelReader");
            } catch {
                Cleanup();
                throw;
            }
        }

        [GlobalCleanup]
        public void Cleanup() {
            foreach (IntPtr pointer in _native) Marshal.FreeHGlobal(pointer);
            _native.Clear();
            _output.Dispose();
        }

        [Benchmark]
        public long OfficeIMOUtf8() {
            _output.SetLength(0);
            ExcelDocument.WriteRows(_output, Enumerable.Range(0, Rows), ["c0", "c1", "c2", "c3"], WriteOfficeIMORow,
                new ExcelTabularWriteOptions { IncludeHeaders = false, IncludeCellReferences = false, UseSharedStrings = false });
            return _output.Length;
        }

        private void WriteOfficeIMORow(ExcelDocument.ExcelTabularRowWriter row, int index) {
            foreach (Utf8Column column in _columns) row.WriteUtf8(column.At(index));
        }

        [Benchmark(Baseline = true)]
        public long ExcelReaderPublicUtf8() {
            _output.SetLength(0);
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(_output, leaveOpen: true)) {
                using XlsxSheetWriter sheet = workbook.AddSheet("S");
                for (int index = 0; index < Rows; index++) {
                    using XlsxRowWriter row = sheet.StartRow();
                    foreach (Utf8Column column in _columns) row.WriteUtf8(column.At(index));
                }
                sheet.End();
                workbook.End();
            }
            return _output.Length;
        }

        private IntPtr Pin(ReadOnlySpan<byte> bytes) {
            IntPtr pointer = Marshal.AllocHGlobal(bytes.Length);
            bytes.CopyTo(new Span<byte>((void*)pointer, bytes.Length));
            _native.Add(pointer);
            return pointer;
        }

        private static string ExpectedValue(int column, int row) => $"customer-{column}-{row % 5000:D5}";

        private static void Validate(byte[] bytes, string engine) {
            using (MemoryStream stream = new MemoryStream(bytes, writable: false)) {
                using ZipArchive package = new ZipArchive(stream, ZipArchiveMode.Read);
                WrittenWorkbookValidation.ValidateEntries(package);
                WrittenWorkbookValidation.ValidateRelationships(package);
            }
            WorkbookScanWorkload workload = new WorkbookScanWorkload();
            workload.Setup(bytes, ComparisonWorkbookFormat.Xlsx, Rows, Columns,
                row => Enumerable.Range(0, Columns).Select(column => (object)ExpectedValue(column, row)).ToArray());
            Console.WriteLine($"Validated {engine} public UTF8 write: rows={Rows}, columns={Columns}, bytes={bytes.Length}, "
                + "preparedInput=unmanagedOffsetsAndUtf8, retainedOutputCapacity=32MiB, headers=false.");
        }

        private readonly record struct Utf8Column(IntPtr Offsets, IntPtr Data) {
            internal ReadOnlySpan<byte> At(int row) {
                int* offsets = (int*)Offsets;
                int start = offsets[row];
                return new ReadOnlySpan<byte>((byte*)Data + start, offsets[row + 1] - start);
            }
        }
    }
}
#endif