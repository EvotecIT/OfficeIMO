using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Measures the same public, plain-tabular write contract for OfficeIMO's two
/// native binary workbook formats. This is a throughput/allocation lane, not a
/// cross-format ranking: XLS and XLSB are different physical contracts.
/// </summary>
[MemoryDiagnoser]
public class ExcelNativeBinaryWriteBenchmarks {
    private BinaryWriteRow[] _rows = null!;

    [Params(2_500, 25_000)]
    public int RowCount { get; set; }

    [Params(ExcelFileFormat.Xls, ExcelFileFormat.Xlsb)]
    public ExcelFileFormat Format { get; set; }

    [GlobalSetup]
    public void Setup() {
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

    private byte[] WriteWorkbook() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Id");
        sheet.CellValue(1, 2, "Region");
        sheet.CellValue(1, 3, "Owner");
        sheet.CellValue(1, 4, "Amount");
        sheet.CellValue(1, 5, "Active");

        for (int index = 0; index < _rows.Length; index++) {
            int row = index + 2;
            BinaryWriteRow value = _rows[index];
            sheet.CellValue(row, 1, value.Id);
            sheet.CellValue(row, 2, value.Region);
            sheet.CellValue(row, 3, value.Owner);
            sheet.CellValue(row, 4, value.Amount);
            sheet.CellValue(row, 5, value.Active);
        }

        return document.ToBytes(Format);
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
