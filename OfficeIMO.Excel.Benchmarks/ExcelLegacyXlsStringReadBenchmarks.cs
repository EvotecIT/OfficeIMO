using System.Data.Common;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Excel.LegacyXls.Read;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Reads BIFF8 string cells through the public XLS data reader: inline LABEL cells in
/// compressed and UTF-16 form, cached formula STRING results, and long shared strings that
/// continue across CONTINUE records. GetChars copies every value in 16-character chunks.
/// Setup proves the XLS fast path opens the workbook and checks every delivered string.
/// This is an OfficeIMO before/after lane.
/// </summary>
[MemoryDiagnoser]
public class ExcelLegacyXlsStringReadBenchmarks {
    private const int ColumnCount = 4;
    private readonly char[] _chunk = new char[16];
    private byte[] _workbook = [];
    private long _expected;

    [Params(5000)]
    public int RowCount { get; set; }

    [Params("Label", "LabelUtf16", "FormulaString", "ContinuedSharedString")]
    public string Cells { get; set; } = "Label";

    [Params("GetString", "GetChars")]
    public string Access { get; set; } = "GetString";

    [GlobalSetup]
    public void Setup() {
        if (Access is not ("GetString" or "GetChars")) throw new ArgumentOutOfRangeException(nameof(Access));
        _workbook = LegacyXlsBenchmarkWorkbook.Create(Cells, RowCount, ColumnCount);
        var options = new ExcelReaderOptions { HasHeaderRow = false };
        using (LegacyXlsTabularWorkbook workbook = LegacyXlsTabularWorkbook.Open(_workbook, options))
        using (DbDataReader direct = workbook.OpenTable(LegacyXlsBenchmarkWorkbook.SheetName, hasHeaderRow: false, options)) {
            if (direct is not LegacyXlsTabularDataReader) {
                throw new InvalidOperationException("The XLS fast path did not open the benchmark workbook.");
            }
            Validate(direct);
        }
        _expected = ReadCore();
    }

    [Benchmark]
    public long Read() {
        long result = ReadCore();
        if (result != _expected) throw new InvalidDataException("The XLS string checksum changed.");
        return result;
    }

    private long ReadCore() {
        using DbDataReader reader = ExcelDocument.OpenReader(_workbook, new ExcelReaderOptions { HasHeaderRow = false });
        bool chunked = Access == "GetChars";
        long checksum = 0;
        int rows = 0;
        while (reader.Read()) {
            for (int column = 0; column < ColumnCount; column++) {
                checksum += chunked ? ConsumeChars(reader, column) : Consume(reader.GetString(column));
            }
            rows++;
        }
        if (rows != RowCount) throw new InvalidDataException("The XLS row count changed.");
        return checksum;
    }

    private static long Consume(string value) => value.Length * 31L + value[value.Length - 1];

    private long ConsumeChars(DbDataReader reader, int column) {
        long offset = 0;
        char last = '\0';
        long read;
        while ((read = reader.GetChars(column, offset, _chunk, 0, _chunk.Length)) > 0) {
            last = _chunk[read - 1];
            offset += read;
        }
        return offset * 31L + last;
    }

    private void Validate(DbDataReader reader) {
        if (reader.FieldCount != ColumnCount) throw new InvalidDataException("The XLS field count differs.");
        int row = 0;
        while (reader.Read()) {
            for (int column = 0; column < ColumnCount; column++) {
                string expected = LegacyXlsBenchmarkWorkbook.Text(Cells, row, column, ColumnCount);
                if (!string.Equals(reader.GetString(column), expected, StringComparison.Ordinal)
                    || ConsumeChars(reader, column) != Consume(expected)) {
                    throw new InvalidDataException($"XLS cell R{row + 1}C{column + 1} differs.");
                }
            }
            row++;
        }
        if (row != RowCount) throw new InvalidDataException("The XLS row count differs.");
    }
}
