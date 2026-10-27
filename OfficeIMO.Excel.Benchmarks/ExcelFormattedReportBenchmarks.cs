using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Benchmarks {
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
            ValidateWorkbook(WriteReport(validate: true));
        }

        [Benchmark]
        public int WriteFormattedReport() => WriteReport().Length;

        /// <summary>Includes validation of the encoded package before committing the destination stream.</summary>
        [Benchmark]
        public int WriteValidatedFormattedReport() => WriteReport(validate: true).Length;

        private byte[] WriteReport(bool validate = false) {
            using MemoryStream stream = new MemoryStream();
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
                    sheet.SetColumnWidth(column, OfficeIMO.Excel.ExcelColumnWidth.Characters(18 + column));
                    sheet.WrapCells(1,RowCount + 1,column);
                }
                for (int row = 2; row <= RowCount + 1; row++) sheet.SetRowHeight(row, OfficeIMO.Excel.ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(42)));
                sheet.SetGridlinesVisible(false);
                sheet.SetOrientation(OfficePageOrientation.Landscape);
                sheet.SetMarginsPreset(ExcelMarginPreset.Narrow);
                sheet.SetPageSetup(fitToWidth: 1, fitToHeight: 0);
                document.Save(stream, validate ? new ExcelSaveOptions { ValidateOpenXml = true } : null);
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
            using MemoryStream stream = new MemoryStream(workbook, writable: false);
            using SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false);
            WorkbookPart workbookPart = package.WorkbookPart ?? throw new InvalidOperationException("The report has no workbook part.");
            Worksheet worksheet = workbookPart.WorksheetParts.Single().Worksheet ?? throw new InvalidOperationException("The report has no worksheet.");
            Row[] rows = worksheet.GetFirstChild<SheetData>()!.Elements<Row>().ToArray();
            if (rows.Length != RowCount + 1) throw new InvalidOperationException($"Expected {RowCount + 1} rows, found {rows.Length}.");
            if (rows.Skip(1).Any(row => row.Height?.Value != 42D || row.CustomHeight?.Value != true)) {
                throw new InvalidOperationException("Every data row must have the fixed 42 point height.");
            }
            Stylesheet styles = workbookPart.WorkbookStylesPart!.Stylesheet!;
            CellFormat[] formats = styles.CellFormats!.Elements<CellFormat>().ToArray();
            DocumentFormat.OpenXml.Spreadsheet.Font[] fonts = styles.Fonts!.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ToArray();
            Fill[] fills = styles.Fills!.Elements<Fill>().ToArray();
            string[] strings = workbookPart.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>()
                .Select(item => item.InnerText).ToArray() ?? [];
            for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++) {
                Cell[] cells = rows[rowIndex].Elements<Cell>().ToArray();
                if (rows[rowIndex].RowIndex?.Value != rowIndex + 1 || cells.Length != Columns) {
                    throw new InvalidOperationException("Report rows and columns must remain complete and ordered.");
                }
                for (int columnIndex = 0; columnIndex < cells.Length; columnIndex++) {
                    Cell cell = cells[columnIndex];
                    string expectedReference = ((char)('A' + columnIndex)).ToString()
                        + (rowIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    if (cell.CellReference?.Value != expectedReference) {
                        throw new InvalidOperationException($"Cell {cell.CellReference} has an unexpected coordinate.");
                    }
                    string actual = cell.DataType?.Value == CellValues.SharedString
                        ? strings[int.Parse(cell.CellValue!.Text, System.Globalization.CultureInfo.InvariantCulture)]
                        : cell.InlineString?.InnerText ?? cell.CellValue?.Text ?? string.Empty;
                    string expected = rowIndex == 0 ? "Header " + (columnIndex + 1)
                        : _values[((rowIndex + 1) * 31 + (columnIndex + 1) * 7) % _values.Length];
                    CellFormat format = formats[(int)(cell.StyleIndex?.Value ?? 0U)];
                    if (actual != expected || format.Alignment?.WrapText?.Value != true) {
                        throw new InvalidOperationException($"Cell {cell.CellReference} lost its value or wrapping.");
                    }
                    if (rowIndex > 0 && columnIndex == 1) {
                        var bold = fonts[(int)format.FontId!.Value].Bold;
                        string expectedFill = "FF" + Fills[((rowIndex + 1) / 50) % Fills.Length].TrimStart('#');
                        if (bold == null || bold.Val?.Value == false
                            || fills[(int)format.FillId!.Value].PatternFill?.ForegroundColor?.Rgb?.Value != expectedFill) {
                            throw new InvalidOperationException($"Status cell {cell.CellReference} lost its font or fill.");
                        }
                    }
                }
            }
            Table table = workbookPart.WorksheetParts.Single().TableDefinitionParts.Single().Table!;
            if (table.Reference?.Value != "A1:L" + (RowCount + 1) || table.AutoFilter?.Reference?.Value != table.Reference?.Value) {
                throw new InvalidOperationException("The report table and its filter must cover every row.");
            }
            if (worksheet.GetFirstChild<PageSetup>()?.Orientation?.Value != OrientationValues.Landscape) {
                throw new InvalidOperationException("The page setup was not saved.");
            }
        }
    }
}
