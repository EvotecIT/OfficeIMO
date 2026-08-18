using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Measures equivalent public, high-throughput tabular write contracts for
/// OfficeIMO's native binary workbook formats and ExcelReader.NET.
/// </summary>
[MemoryDiagnoser]
public class ExcelNativeBinaryWriteBenchmarks {
    private BinaryWriteRow[] _rows = null!;

    [Params(2_500, 25_000)]
    public int RowCount { get; set; }

    // ExcelReader.NET 2.1.2 omits the required BrtWsDim record from XLSB.
    // Its XLS writer also omits the BIFF8 Index/Row/DBCell cell-table structure,
    // so this XLS comparison is a diagnostic throughput lane rather than a
    // format-conformance-equivalent result.
    [Params(ExcelFileFormat.Xls)]
    public ExcelFileFormat Format { get; set; }

    [GlobalSetup]
    public void Setup() {
        SetupOfficeIMOOnly();
        byte[] excelReaderWorkbook = WriteExcelReaderWorkbook();
        Validate(excelReaderWorkbook);
    }

    internal void SetupOfficeIMOOnly() {
        _rows = new BinaryWriteRow[RowCount];
        for (int index = 0; index < _rows.Length; index++) {
            _rows[index] = new BinaryWriteRow(
                index + 1,
                "Region " + (index % 8),
                "Owner " + (index % 32),
                Math.Round(100d + ((index * 17.25d) % 9_000d), 2),
                (index & 1) == 0);
        }

        byte[] workbook = WriteWorkbook();
        Validate(workbook);
    }

    [Benchmark]
    public int OfficeIMO_PublicTabularWrite() => WriteWorkbook().Length;

    [Benchmark]
    public int ExcelReaderNet_DiagnosticWrite() => WriteExcelReaderWorkbook().Length;

    private byte[] WriteWorkbook() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.InsertObjects(
            _rows,
            ("Id", static row => row.Id),
            ("Region", static row => row.Region),
            ("Owner", static row => row.Owner),
            ("Amount", static row => row.Amount),
            ("Active", static row => row.Active));

        byte[] workbook = document.ToBytes(Format);
        if (document.LastSaveDiagnostics.Writer != ExcelSavePackageWriter.NativeBinaryDirectPackage) {
            throw new InvalidOperationException(
                $"OfficeIMO {Format} benchmark did not use the native direct tabular writer: "
                + document.LastSaveDiagnostics.FastPackageSkipReason);
        }

        return workbook;
    }

    private byte[] WriteExcelReaderWorkbook() => Format switch {
        ExcelFileFormat.Xls => WriteExcelReaderXlsWorkbook(),
        ExcelFileFormat.Xlsb => WriteExcelReaderXlsbWorkbook(),
        _ => throw new InvalidOperationException($"ExcelReader.NET binary benchmark does not support {Format}.")
    };

    private byte[] WriteExcelReaderXlsWorkbook() {
        using var stream = new MemoryStream();
        using XlsWorkbookWriter workbook = XlsWorkbookWriter.Create(stream, leaveOpen: true);
        workbook.Start();
        using (XlsSheetWriter sheet = workbook.AddSheet("Data")) {
            sheet.Start();
            using (XlsRowWriter header = sheet.StartRow()) {
                WriteHeaders(header);
            }

            foreach (BinaryWriteRow value in _rows) {
                using XlsRowWriter row = sheet.StartRow();
                WriteRow(row, value);
            }
            sheet.End();
        }
        workbook.End();
        return stream.ToArray();
    }

    private byte[] WriteExcelReaderXlsbWorkbook() {
        using var stream = new MemoryStream();
        using XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(stream, leaveOpen: true);
        workbook.Start();
        using (XlsbSheetWriter sheet = workbook.AddSheet("Data")) {
            sheet.Start();
            using (XlsbRowWriter header = sheet.StartRow()) {
                WriteHeaders(header);
            }

            foreach (BinaryWriteRow value in _rows) {
                using XlsbRowWriter row = sheet.StartRow();
                WriteRow(row, value);
            }
            sheet.End();
        }
        workbook.End();
        return stream.ToArray();
    }

    private static void WriteHeaders<TRow>(TRow row)
        where TRow : IRowWriter {
        row.Write("Id");
        row.Write("Region");
        row.Write("Owner");
        row.Write("Amount");
        row.Write("Active");
    }

    private static void WriteRow<TRow>(TRow row, BinaryWriteRow value)
        where TRow : IRowWriter {
        row.Write(value.Id);
        row.Write(value.Region);
        row.Write(value.Owner);
        row.Write(value.Amount);
        row.Write(value.Active);
    }

    private void Validate(byte[] workbook) {
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(workbook);
        string[] headers = ["Id", "Region", "Owner", "Amount", "Active"];
        if (reader.FieldCount != headers.Length) {
            throw new InvalidDataException($"{Format} exposed {reader.FieldCount} fields instead of {headers.Length}.");
        }

        for (int column = 0; column < headers.Length; column++) {
            if (!string.Equals(reader.GetName(column), headers[column], StringComparison.Ordinal)) {
                throw new InvalidDataException($"{Format} header {column + 1} did not round-trip.");
            }
        }

        int rowIndex = 0;
        while (reader.Read()) {
            if (rowIndex >= _rows.Length) {
                throw new InvalidDataException($"{Format} emitted extra rows.");
            }

            BinaryWriteRow expected = _rows[rowIndex];
            if (reader.GetInt32(0) != expected.Id
                || !string.Equals(reader.GetString(1), expected.Region, StringComparison.Ordinal)
                || !string.Equals(reader.GetString(2), expected.Owner, StringComparison.Ordinal)
                || reader.GetDouble(3) != expected.Amount
                || reader.GetBoolean(4) != expected.Active) {
                throw new InvalidDataException($"{Format} row {rowIndex + 2} did not round-trip.");
            }

            rowIndex++;
        }

        if (rowIndex != _rows.Length || reader.NextResult()) {
            throw new InvalidDataException($"{Format} did not round-trip the exact single-sheet row set.");
        }
    }

    private readonly record struct BinaryWriteRow(int Id, string Region, string Owner, double Amount, bool Active);
}

/// <summary>
/// Measures OfficeIMO's validated native binary write paths without ranking
/// structurally invalid competitor output.
/// </summary>
[MemoryDiagnoser]
public class OfficeNativeBinaryWriteBenchmarks {
    private ExcelNativeBinaryWriteBenchmarks _scenario = null!;

    [Params(2_500, 25_000)]
    public int RowCount { get; set; }

    [Params(ExcelFileFormat.Xls, ExcelFileFormat.Xlsb)]
    public ExcelFileFormat Format { get; set; }

    [GlobalSetup]
    public void Setup() {
        _scenario = new ExcelNativeBinaryWriteBenchmarks {
            RowCount = RowCount,
            Format = Format
        };
        _scenario.SetupOfficeIMOOnly();
    }

    [Benchmark]
    public int OfficeIMO_ValidatedNativeBinaryWrite() =>
        _scenario.OfficeIMO_PublicTabularWrite();
}
