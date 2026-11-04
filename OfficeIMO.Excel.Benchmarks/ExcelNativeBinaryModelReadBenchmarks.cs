using BenchmarkDotNet.Attributes;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.Xlsb;
using XlsbCell = OfficeIMO.Excel.Xlsb.Model.XlsbCell;
using XlsbWorkbook = OfficeIMO.Excel.Xlsb.Model.XlsbWorkbook;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Loads native XLS and XLSB workbooks into OfficeIMO's full-fidelity workbook models.
/// </summary>
/// <remarks>
/// The tabular readers measured elsewhere stream cells without building a model. This
/// scenario measures the compound-file, BIFF8 and BIFF12 record paths behind opening,
/// editing and converting a binary workbook. Every load is validated against the
/// generated cell count and numeric checksum.
/// </remarks>
[MemoryDiagnoser]
public class ExcelNativeBinaryModelReadBenchmarks {
    private const int ColumnCount = 5;
    private byte[] _workbook = Array.Empty<byte>();
    private double _expectedNumericChecksum;

    [Params(ExcelFileFormat.Xls, ExcelFileFormat.Xlsb)]
    public ExcelFileFormat Format { get; set; }

    [Params(20_000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        var rows = new ModelReadRow[RowCount];
        for (int index = 0; index < rows.Length; index++) {
            rows[index] = new ModelReadRow(
                index + 1,
                "Region " + (index % 8),
                "Owner " + (index % 32),
                Math.Round(100d + ((index * 17.25d) % 9_000d), 2),
                (index & 1) == 0);
            _expectedNumericChecksum += rows[index].Id + rows[index].Amount;
        }

        using (ExcelDocument document = ExcelDocument.Create()) {
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.WriteData(
                rows,
                ("Id", static _ => _.Id),
                ("Region", static _ => _.Region),
                ("Owner", static _ => _.Owner),
                ("Amount", static _ => _.Amount),
                ("Active", static _ => _.Active));
            _workbook = document.ToBytes(Format);
        }

        LoadModel();
    }

    [Benchmark]
    public int LoadModel() {
        (int cellCount, double numericChecksum) = Format == ExcelFileFormat.Xls ? LoadXls() : LoadXlsb();
        int expectedCells = (RowCount + 1) * ColumnCount;
        if (cellCount != expectedCells) {
            throw new InvalidDataException($"{Format} model load produced {cellCount} cells instead of {expectedCells}.");
        }

        if (Math.Abs(numericChecksum - _expectedNumericChecksum) > 0.001d) {
            throw new InvalidDataException($"{Format} model load produced numeric checksum {numericChecksum} instead of {_expectedNumericChecksum}.");
        }

        return cellCount;
    }

    private (int CellCount, double NumericChecksum) LoadXls() {
        LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(_workbook);
        IReadOnlyList<LegacyXlsCell> cells = workbook.Worksheets[0].Cells;
        double checksum = 0d;
        foreach (LegacyXlsCell cell in cells) {
            if (cell.Value is double number) checksum += number;
        }

        return (cells.Count, checksum);
    }

    private (int CellCount, double NumericChecksum) LoadXlsb() {
        XlsbWorkbook workbook = XlsbWorkbookReader.Load(_workbook);
        IReadOnlyList<XlsbCell> cells = workbook.Worksheets[0].Cells;
        double checksum = 0d;
        foreach (XlsbCell cell in cells) {
            if (cell.Value is double number) checksum += number;
        }

        return (cells.Count, checksum);
    }

    private sealed record ModelReadRow(int Id, string Region, string Owner, double Amount, bool Active);
}
