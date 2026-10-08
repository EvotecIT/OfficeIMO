using System.Data;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Materializes mixed object columns with sparse and repeated physical rows.</summary>
[MemoryDiagnoser]
public class ExcelDataTableRowShapeBenchmarks {
    private string _path = string.Empty;
    private string _range = string.Empty;

    [Params(1000)] public int RowCount { get; set; }
    [Params(8, 65)] public int ColumnCount { get; set; }
    [Params(false, true)] public bool Sparse { get; set; }
    [Params(false, true)] public bool RepeatedRow { get; set; }
    [Params(false, true)] public bool ReverseRows { get; set; }

    [GlobalSetup]
    public void Setup() {
        if (RowCount < 2) throw new ArgumentOutOfRangeException(nameof(RowCount));
        string lastColumn = ColumnCount switch {
            8 => "H", 65 => "BM", _ => throw new ArgumentOutOfRangeException(nameof(ColumnCount))
        };
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_DATA") ?? Path.GetTempPath();
        Directory.CreateDirectory(root);
        _path = Path.Combine(root, $"datatable-row-shapes-{Guid.NewGuid():N}.xlsx");
        _range = $"A1:{lastColumn}{RowCount}";
        try {
            using (var document = ExcelDocument.Create(_path)) {
                document.AddWorksheet("Data");
                document.Save();
            }
            using (var package = ZipFile.Open(_path, ZipArchiveMode.Update)) {
                const string entryName = "xl/worksheets/sheet1.xml";
                package.GetEntry(entryName)!.Delete();
                using var writer = new StreamWriter(package.CreateEntry(entryName, CompressionLevel.Fastest).Open(), new UTF8Encoding(false));
                writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
                for (int index = 1; index <= RowCount; index++) {
                    int row = ReverseRows ? RowCount - index + 1 : index;
                    if (Sparse && row % 5 == 0) continue;
                    WriteRow(writer, row, replacement: false);
                }
                if (RepeatedRow) WriteRow(writer, 1, replacement: true);
                writer.Write("</sheetData></worksheet>");
            }
            using var table = ReadTable();
            for (int column = 0; column < ColumnCount; column++) {
                if (table.Columns[column].DataType != typeof(object)
                    || table.Columns[column].ColumnName != $"Column{column + 1}")
                    throw new InvalidDataException("Object-column schema differs.");
            }
            for (int row = 1; row <= RowCount; row++) {
                for (int column = 1; column <= ColumnCount; column++) {
                    object actual = table.Rows[row - 1][column - 1];
                    object expected = ExpectedValue(row, column);
                    if (actual.GetType() != expected.GetType() || !Equals(actual, expected))
                        throw new InvalidDataException($"Row shape value differs at row {row}, column {column}.");
                }
            }
        } catch {
            Cleanup();
            throw;
        }
    }

    [Benchmark]
    public int Read() {
        using var table = ReadTable();
        return table.Rows.Count;
    }

    [GlobalCleanup]
    public void Cleanup() {
        if (_path.Length != 0) File.Delete(_path);
    }

    private DataTable ReadTable() {
        using var owner = ExcelDocumentReader.Open(_path, new ExcelReadOptions { InferDataTableColumnTypes = false });
        DataTable table = owner.GetSheet("Data").ReadRangeAsDataTable(_range, headersInFirstRow: false);
        if (table.Rows.Count == RowCount && table.Columns.Count == ColumnCount) return table;
        table.Dispose();
        throw new InvalidDataException("DataTable shape differs.");
    }

    private object ExpectedValue(int row, int column) {
        if (RepeatedRow && row == 1) return $"replacement-{column}";
        if (Sparse && (row % 5 == 0 || column % 4 == 0)) return DBNull.Value;
        return (column % 3) switch {
            0 => $"row-{row}-column-{column}",
            1 => (double)(row * 100 + column),
            _ => row % 2 == 0
        };
    }

    private void WriteRow(TextWriter writer, int row, bool replacement) {
        writer.Write("<row r=\"");
        writer.Write(row.ToString(CultureInfo.InvariantCulture));
        writer.Write("\">");
        for (int column = 1; column <= ColumnCount; column++) {
            if (!replacement && Sparse && column % 4 == 0) {
                writer.Write("<c/>");
            } else if (replacement || column % 3 == 0) {
                writer.Write("<c t=\"str\"><v>");
                writer.Write(replacement ? $"replacement-{column}" : $"row-{row}-column-{column}");
                writer.Write("</v></c>");
            } else if (column % 3 == 1) {
                writer.Write("<c><v>");
                writer.Write((row * 100 + column).ToString(CultureInfo.InvariantCulture));
                writer.Write("</v></c>");
            } else {
                writer.Write(row % 2 == 0 ? "<c t=\"b\"><v>1</v></c>" : "<c t=\"b\"><v>0</v></c>");
            }
        }
        writer.Write("</row>");
    }
}
