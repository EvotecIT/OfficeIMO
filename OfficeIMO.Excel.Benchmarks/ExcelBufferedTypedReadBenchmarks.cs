using System.Globalization;
using System.IO.Compression;
using System.Text;
using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Reads a narrow typed projection from ordered and reversed wide worksheets.</summary>
[MemoryDiagnoser]
public class ExcelBufferedTypedReadBenchmarks {
    private string _path = string.Empty;
    private ExcelDocumentReader? _heldOwner;
    private IEnumerator<ProjectedRow>? _heldRows;

    [Params(1000, 5000)] public int RowCount { get; set; }
    [Params(4, 65)] public int ColumnCount { get; set; }
    [Params(false, true)] public bool ReverseRows { get; set; }
    [Params(false, true)] public bool Utf16 { get; set; }

    [GlobalSetup]
    public void Setup() {
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_DATA") ?? Path.GetTempPath();
        Directory.CreateDirectory(root);
        _path = Path.Combine(root, $"buffered-typed-{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(_path)) {
                document.AddWorksheet("Data");
                document.Save();
            }
            using (var package = ZipFile.Open(_path, ZipArchiveMode.Update)) {
                const string part = "xl/worksheets/sheet1.xml";
                package.GetEntry(part)!.Delete();
                using var writer = new StreamWriter(package.CreateEntry(part, CompressionLevel.Fastest).Open(),
                    Utf16 ? new UnicodeEncoding(false, true) : new UTF8Encoding(false));
                writer.Write(Utf16 ? "<?xml version=\"1.0\" encoding=\"utf-16\"?>" : "<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1:");
                writer.Write(ColumnCount == 65 ? "BM" : "D");
                writer.Write((RowCount + 1).ToString(CultureInfo.InvariantCulture));
                writer.Write("\"/><sheetData>");
                for (int index = 1; index <= RowCount + 1; index++) {
                    int row = ReverseRows ? RowCount + 2 - index : index;
                    writer.Write("<row r=\"");
                    writer.Write(row.ToString(CultureInfo.InvariantCulture));
                    writer.Write("\">");
                    if (row == 1) {
                        WriteText(writer, "Id");
                        WriteText(writer, "Name");
                    } else {
                        writer.Write("<c><v>");
                        writer.Write((row - 1).ToString(CultureInfo.InvariantCulture));
                        writer.Write("</v></c>");
                        WriteText(writer, Name(row - 1));
                    }
                    for (int column = 3; column <= ColumnCount; column++)
                        WriteText(writer, $"ignored-cell-{row:D6}-{column:D2}-payload");
                    writer.Write("</row>");
                }
                writer.Write("</sheetData></worksheet>");
            }
            ReadAll(validate: true);
            foreach (var mode in new[] { ExcelExecutionMode.Automatic, ExcelExecutionMode.Sequential, ExcelExecutionMode.Parallel })
                ReadAll(validate: true, mode);
        } catch {
            Cleanup();
            throw;
        }
    }

    [Benchmark]
    public long Read() => ReadAll(validate: false);

    [Benchmark]
    public long ReadAutomatic() => ReadAll(validate: false, ExcelExecutionMode.Automatic);

    [Benchmark]
    public long ReadSequential() => ReadAll(validate: false, ExcelExecutionMode.Sequential);

    [Benchmark]
    public long ReadParallel() => ReadAll(validate: false, ExcelExecutionMode.Parallel);

    [Benchmark]
    public int FirstRow() {
        using var owner = ExcelDocumentReader.Open(_path);
        using var rows = owner.GetSheet("Data").ReadObjectsStream<ProjectedRow>($"A1:B{RowCount + 1}").GetEnumerator();
        if (!rows.MoveNext() || rows.Current.Id != 1 || rows.Current.Name != Name(1))
            throw new InvalidDataException("First projected row differs.");
        return rows.Current.Id;
    }

    /// <summary>Holds the enumerator at its first row for an opt-in retained-memory measurement.</summary>
    public int HoldFirstRow() {
        ReleaseHeldRow();
        _heldOwner = ExcelDocumentReader.Open(_path);
        try {
            _heldRows = _heldOwner.GetSheet("Data").ReadObjectsStream<ProjectedRow>($"A1:B{RowCount + 1}").GetEnumerator();
            if (!_heldRows.MoveNext() || _heldRows.Current.Id != 1 || _heldRows.Current.Name != Name(1))
                throw new InvalidDataException("Held projected row differs.");
            return _heldRows.Current.Id;
        } catch {
            ReleaseHeldRow();
            throw;
        }
    }

    /// <summary>Disposes the reader and enumerator retained by the memory probe.</summary>
    public void ReleaseHeldRow() {
        try { _heldRows?.Dispose(); }
        finally {
            _heldRows = null;
            _heldOwner?.Dispose();
            _heldOwner = null;
        }
    }

    [GlobalCleanup]
    public void Cleanup() {
        ReleaseHeldRow();
        if (_path.Length != 0) File.Delete(_path);
    }

    private long ReadAll(bool validate, ExcelExecutionMode? mode = null) {
        using var owner = ExcelDocumentReader.Open(_path);
        int count = 0;
        long observation = 0;
        var sheet = owner.GetSheet("Data");
        string range = $"A1:B{RowCount + 1}";
        var rows = mode == null ? sheet.ReadObjectsStream<ProjectedRow>(range) : sheet.ReadObjects<ProjectedRow>(range, mode);
        foreach (var row in rows) {
            count++;
            if (validate && (row.Id != count || row.Name != Name(count)))
                throw new InvalidDataException($"Projected row {count} differs.");
            observation += row.Id + row.Name!.Length;
        }
        long expected = (long)RowCount * (RowCount + 1) / 2 + (long)RowCount * Name(1).Length;
        if (count != RowCount || observation != expected)
            throw new InvalidDataException("Complete projected observation differs.");
        return observation;
    }

    private static string Name(int id) => "row-" + id.ToString("D6", CultureInfo.InvariantCulture);
    private static void WriteText(TextWriter writer, string value) {
        writer.Write("<c t=\"str\"><v>");
        writer.Write(value);
        writer.Write("</v></c>");
    }

    /// <summary>The two projected columns; the remaining physical columns are deliberately ignored.</summary>
    public sealed class ProjectedRow {
        public int Id { get; set; }
        public string? Name { get; set; }
    }
}
