using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Measures a formatted report sheet written cell by cell: text rows with one styled status cell per row, a table,
/// a frozen header, column widths, wrapped columns, a fixed height for every data row and page setup, then a save.
/// Layout setters called once per row or column must stay proportional to the rows written.
/// </summary>
[MemoryDiagnoser]
public class ExcelFormattedReportBenchmarks {
    private const int Columns = 12; // A:L
    private static readonly string[] Fills = ["#F8D7DA", "#E6F4EA", "#FFF3CD", "#EEE8FF"];
    private string[] _values = [];

    [Params(5_000, 25_000)]
    public int RowCount { get; set; }

    /// <summary>Whether the per-row writes and styling run inside <see cref="ExcelSheet.Batch"/>.</summary>
    [Params(false, true)]
    public bool StyleInBatch { get; set; }

    [GlobalSetup]
    public void Setup() {
        _values = Enumerable.Range(0, 4_000).Select(index => "Value " + index + " " + new string('x', 20 + index % 60)).ToArray();
        ValidateWorkbook(WriteReport());
    }

    [Benchmark]
    public int WriteFormattedReport() => WriteReport().Length;

    private byte[] WriteReport() {
        using var stream = new MemoryStream();
        using (ExcelDocument document = ExcelDocument.Create(stream)) {
            ExcelSheet sheet = document.AddWorksheet("Report");
            for (int column = 1; column <= Columns; column++) {
                sheet.CellValue(1, column, "Header " + column);
                sheet.CellAt(1, column).HeaderStyle();
            }
            if (StyleInBatch) sheet.Batch(WriteRows); else WriteRows(sheet);
            sheet.Range("A1:L" + (RowCount + 1)).CreateTable("Report", hasHeader: true, ExcelTableStyle.TableStyleMedium2, includeAutoFilter: true);
            sheet.Freeze(1, 0);
            for (int column = 1; column <= Columns; column++) {
                sheet.SetColumnWidth(column, 18 + column);
                sheet.WrapCells(1, RowCount + 1, column);
            }
            for (int row = 2; row <= RowCount + 1; row++) sheet.SetRowHeight(row, 42);
            sheet.SetGridlinesVisible(false);
            sheet.SetOrientation(OfficePageOrientation.Landscape);
            sheet.SetMarginsPreset(ExcelMarginPreset.Narrow);
            sheet.SetPageSetup(fitToWidth: 1, fitToHeight: 0);
            document.Save(stream);
        }
        return stream.ToArray();
    }

    private void WriteRows(ExcelSheet sheet) {
        for (int row = 2; row <= RowCount + 1; row++) {
            for (int column = 1; column <= Columns; column++) {
                sheet.CellValue(row, column, _values[(row * 31 + column * 7) % _values.Length]);
            }
            sheet.CellAt(row, 2).SetFillColor(Fills[(row / 50) % Fills.Length]).SetBold(true);
        }
    }

    private void ValidateWorkbook(byte[] workbook) {
        using var stream = new MemoryStream(workbook, writable: false);
        using SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false);
        WorkbookPart workbookPart = package.WorkbookPart ?? throw new InvalidOperationException("The report has no workbook part.");
        Worksheet worksheet = workbookPart.WorksheetParts.Single().Worksheet ?? throw new InvalidOperationException("The report has no worksheet.");
        Row[] rows = worksheet.GetFirstChild<SheetData>()!.Elements<Row>().ToArray();
        if (rows.Length != RowCount + 1) throw new InvalidOperationException($"Expected {RowCount + 1} rows, found {rows.Length}.");
        if (rows.Skip(1).Any(row => row.Height?.Value != 42D || row.CustomHeight?.Value != true)) {
            throw new InvalidOperationException("Every data row must have the fixed 42 point height.");
        }
        Stylesheet styles = workbookPart.WorkbookStylesPart!.Stylesheet!;
        if (!styles.CellFormats!.Elements<CellFormat>().Any(format => format.Alignment?.WrapText?.Value == true)) {
            throw new InvalidOperationException("The wrapped columns have no wrap-text format.");
        }
        if (worksheet.GetFirstChild<PageSetup>()?.Orientation?.Value != OrientationValues.Landscape) {
            throw new InvalidOperationException("The page setup was not saved.");
        }
    }
}
