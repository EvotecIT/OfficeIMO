#if OFFICEIMO_BENCHMARK_NEW_APIS
using System.Buffers;
using System.Data.Common;
using System.Runtime.InteropServices;
using System.Text;
using BenchmarkDotNet.Attributes;
using OfficeIMO.CSV;
using PeerCsvWorkbookWriter = ExcelReader.Core.Writer.Csv.CsvWorkbookWriter;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>
/// Public UTF-8 row APIs over the upstream unmanaged offsets and string-byte workload.
/// Caller row orchestration is measured. The peer's internal Native WriteApi remains a
/// separate source-only diagnostic and is not called by this public companion.
/// </summary>
// Input provenance: ExcelReader ca5b50f99e8ef57ab476f0a2bc8043558d58b28d,
// tests/ExcelReader.Benchmarks/Write/WritePathBenchmark.cs, NativeStrings_Csv.
[MemoryDiagnoser]
[BenchmarkCategory("CsvWriteUtf8Public")]
public unsafe class CsvUtf8WriterBenchmarks {
    private const int Rows = 100_000, Columns = 4;
    private static readonly string[] Headers = ["c0", "c1", "c2", "c3"];
    private readonly MemoryStream _output = new(32 * 1024 * 1024);
    private readonly ReadOnlyMemory<byte>?[] _row = new ReadOnlyMemory<byte>?[Columns];
    private Utf8Column[] _columns = [];
    private long _qualifiedLength;

    [GlobalSetup]
    public void Setup() {
        BenchmarkInput.WriteDescription();
        try {
            _columns = new Utf8Column[Columns];
            for (int column = 0; column < Columns; column++) {
                int[] offsets = new int[Rows + 1];
                using var data = new MemoryStream();
                for (int row = 0; row < Rows; row++) {
                    data.Write(Encoding.UTF8.GetBytes(ExpectedValue(column, row)));
                    offsets[row + 1] = checked((int)data.Length);
                }
                byte[] values = data.ToArray();
                BenchmarkInput.WriteFixtureIdentity($"write/Utf8/Csv/rows={Rows}/column={column}/offsets",
                    MemoryMarshal.AsBytes(offsets.AsSpan()));
                BenchmarkInput.WriteFixtureIdentity($"write/Utf8/Csv/rows={Rows}/column={column}/values", values);
                _columns[column] = new Utf8Column(offsets, values);
            }
            _qualifiedLength = Rows * (Columns * Encoding.UTF8.GetByteCount(ExpectedValue(0, 0)) + Columns - 1 + 2);
            ExcelReaderPublicUtf8();
            ValidateOutput("ExcelReader");
            OfficeIMOPublicUtf8Row();
            ValidateOutput("OfficeIMO");
            ValidateEscapingAndNulls();
        } catch {
            Cleanup();
            throw;
        }
    }

    [GlobalCleanup]
    public void Cleanup() {
        Array.Clear(_row);
        foreach (Utf8Column? column in _columns) column?.Dispose();
        _columns = [];
        _output.Dispose();
    }

    [Benchmark]
    public long OfficeIMOPublicUtf8Row() {
        _output.SetLength(0);
        using (var writer = CsvRowWriter.CreateStream(_output,
                   new CsvSaveOptions { IncludeHeader = false, NewLine = "\r\n" }, leaveOpen: true)) {
            for (int row = 0; row < Rows; row++) {
                for (int column = 0; column < Columns; column++) _row[column] = _columns[column].At(row);
                if (row == 0) writer.WriteUtf8Row(Headers, _row);
                else writer.WriteUtf8Row(_row);
            }
        }
        return CheckLength();
    }

    [Benchmark(Baseline = true)]
    public long ExcelReaderPublicUtf8() {
        _output.SetLength(0);
        using (var workbook = PeerCsvWorkbookWriter.Create(_output, leaveOpen: true)) {
            using var sheet = workbook.AddSheet("S");
            for (int row = 0; row < Rows; row++) {
                using var outputRow = sheet.StartRow();
                foreach (Utf8Column column in _columns) outputRow.WriteUtf8(column.AtSpan(row));
            }
            sheet.End();
            workbook.End();
        }
        return CheckLength();
    }

    private long CheckLength() => _output.Length == _qualifiedLength
        ? _output.Length : throw new InvalidDataException("Qualified UTF-8 CSV output length changed.");

    private void ValidateOutput(string engine) {
        ReadOnlySpan<byte> remaining = _output.GetBuffer().AsSpan(0, checked((int)_output.Length));
        for (int row = 0; row < Rows; row++) {
            for (int column = 0; column < Columns; column++) {
                ReadOnlySpan<byte> expected = _columns[column].AtSpan(row);
                if (!remaining.StartsWith(expected)) throw new InvalidDataException($"{engine} changed UTF-8 bytes at {row}/{column}.");
                remaining = remaining[expected.Length..];
                ReadOnlySpan<byte> separator = column + 1 < Columns ? ","u8 : "\r\n"u8;
                if (!remaining.StartsWith(separator)) throw new InvalidDataException($"{engine} changed a CSV record boundary.");
                remaining = remaining[separator.Length..];
            }
        }
        if (!remaining.IsEmpty) throw new InvalidDataException($"{engine} wrote unexpected trailing bytes.");
        ValidateFields(_output, Rows, static (row, column) => ExpectedValue(column, row));
        Console.WriteLine($"CSV public UTF8 writer={engine}, rows={Rows}, columns={Columns}, bytes={_output.Length}, "
            + $"sha256={Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(_output.GetBuffer().AsSpan(0, checked((int)_output.Length))))}, "
            + "preparedInput=unmanagedOffsetsAndUtf8, retainedOutputCapacity=32MiB, headers=false, newline=CRLF; every byte and field validated.");
    }

    private static void ValidateEscapingAndNulls() {
        string?[][] fields = [["a,b", "a\"b", "a\rb\nc", "Zażółć 😀"], [null, "", "=SUM(A1)", " plain "]];
        using var office = new MemoryStream();
        using (var writer = CsvRowWriter.CreateStream(office,
                   new CsvSaveOptions { IncludeHeader = false, NewLine = "\r\n" }, leaveOpen: true)) {
            foreach (string?[] row in fields) writer.WriteUtf8Row(Headers,
                row.Select(static value => value is null ? (ReadOnlyMemory<byte>?)null : Encoding.UTF8.GetBytes(value)).ToArray());
        }
        using var peer = new MemoryStream();
        using (var workbook = PeerCsvWorkbookWriter.Create(peer, leaveOpen: true)) {
            using var sheet = workbook.AddSheet("S");
            foreach (string?[] row in fields) {
                using var outputRow = sheet.StartRow();
                foreach (string? value in row) {
                    if (value is null) outputRow.Skip(1);
                    else outputRow.WriteUtf8(Encoding.UTF8.GetBytes(value));
                }
            }
            sheet.End();
            workbook.End();
        }
        byte[] expected = Encoding.UTF8.GetBytes("\"a,b\",\"a\"\"b\",\"a\rb\nc\",Zażółć 😀\r\n,,=SUM(A1), plain \r\n");
        if (!office.ToArray().AsSpan().SequenceEqual(expected) || !peer.ToArray().AsSpan().SequenceEqual(expected))
            throw new InvalidDataException("Public UTF-8 CSV quoting, Unicode, null, empty or formula-preservation output changed.");
        ValidateFields(office, fields.Length, (row, column) => fields[row][column] ?? "");
        ValidateFields(peer, fields.Length, (row, column) => fields[row][column] ?? "");
        using var marked = new MemoryStream();
        using (var writer = CsvRowWriter.CreateStream(marked,
                   new CsvSaveOptions { IncludeHeader = false, NewLine = "\r\n", NullValue = "NULL" }, leaveOpen: true))
            writer.WriteUtf8Row(Headers, [null, ReadOnlyMemory<byte>.Empty, "ok"u8.ToArray(), "last"u8.ToArray()]);
        if (!marked.ToArray().AsSpan().SequenceEqual("NULL,,ok,last\r\n"u8))
            throw new InvalidDataException("The CSV public UTF-8 API conflated a null marker with empty text.");
        Console.WriteLine("CSV public UTF8 writer escaping/Unicode/null/empty/formula-preservation qualification passed.");
    }

    private static void ValidateFields(MemoryStream stream, int rows, Func<int, int, string> expected) {
        stream.Position = 0;
        using (var text = new StreamReader(stream, Encoding.UTF8, true, 1024, leaveOpen: true))
        using (var reader = global::Sylvan.Data.Csv.CsvDataReader.Create(text,
                   new global::Sylvan.Data.Csv.CsvDataReaderOptions { HasHeaders = false }))
            ValidateFields(reader, rows, expected);
        stream.Position = 0;
        using (DbDataReader reader = CsvDocument.OpenDataReader(stream, new CsvLoadOptions { HasHeaderRow = false }))
            ValidateFields(reader, rows, expected);
    }

    private static void ValidateFields(DbDataReader reader, int rows, Func<int, int, string> expected) {
        if (reader.FieldCount != Columns) throw new InvalidDataException("UTF-8 CSV column count changed.");
        for (int row = 0; row < rows; row++) {
            if (!reader.Read()) throw new InvalidDataException("UTF-8 CSV row was omitted.");
            for (int column = 0; column < Columns; column++) {
                if (reader.GetString(column) != expected(row, column))
                    throw new InvalidDataException($"UTF-8 CSV field changed at {row}/{column}.");
            }
        }
        if (reader.Read()) throw new InvalidDataException("UTF-8 CSV wrote extra rows.");
    }

    private static string ExpectedValue(int column, int row) => $"customer-{column}-{row % 5000:D5}";

    private sealed class Utf8Column : IDisposable {
        private readonly IntPtr _offsets;
        private readonly UnmanagedUtf8Memory _data;
        internal Utf8Column(int[] offsets, byte[] data) {
            _offsets = Marshal.AllocHGlobal(checked(offsets.Length * sizeof(int)));
            try {
                offsets.AsSpan().CopyTo(new Span<int>((void*)_offsets, offsets.Length));
                _data = new UnmanagedUtf8Memory(data);
            } catch { Marshal.FreeHGlobal(_offsets); throw; }
        }
        internal ReadOnlyMemory<byte> At(int row) {
            int* offsets = (int*)_offsets;
            return _data.Memory.Slice(offsets[row], offsets[row + 1] - offsets[row]);
        }
        internal ReadOnlySpan<byte> AtSpan(int row) {
            int* offsets = (int*)_offsets;
            return new ReadOnlySpan<byte>((byte*)_data.Pointer + offsets[row], offsets[row + 1] - offsets[row]);
        }
        public void Dispose() {
            Marshal.FreeHGlobal(_offsets);
            ((IDisposable)_data).Dispose();
        }
    }

    // Fixture ownership only: projects the same unmanaged UTF-8 allocation as memory
    // for the public row API. Per-row slicing and array population remain measured.
    private sealed class UnmanagedUtf8Memory : MemoryManager<byte> {
        private IntPtr _data;
        private readonly int _length;
        internal IntPtr Pointer => _data;
        internal UnmanagedUtf8Memory(byte[] bytes) {
            _length = bytes.Length;
            _data = Marshal.AllocHGlobal(_length);
            bytes.AsSpan().CopyTo(new Span<byte>((void*)_data, _length));
        }
        public override Span<byte> GetSpan() {
            ObjectDisposedException.ThrowIf(_data == IntPtr.Zero, this);
            return new Span<byte>((void*)_data, _length);
        }
        public override MemoryHandle Pin(int elementIndex = 0) {
            if ((uint)elementIndex > (uint)_length) throw new ArgumentOutOfRangeException(nameof(elementIndex));
            ObjectDisposedException.ThrowIf(_data == IntPtr.Zero, this);
            return new MemoryHandle((byte*)_data + elementIndex);
        }
        public override void Unpin() { }
        protected override void Dispose(bool disposing) {
            Marshal.FreeHGlobal(_data);
            _data = IntPtr.Zero;
        }
    }
}
#endif
