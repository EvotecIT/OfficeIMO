using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using SixLabors.Fonts;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests for autofitting columns and rows.
    /// </summary>
    public partial class Excel {
        [Fact]
        public void Test_AutoFitColumnsAndRows() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Long piece of text");
                sheet.CellValue(2, 1, "Second line\nwith newline");
                sheet.CellValue(3, 1, "Line1\nLine2\nLine3");
                sheet.AutoFitColumns();
                sheet.AutoFitRows();
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                var columns = wsPart.Worksheet.GetFirstChild<Columns>();
                Assert.NotNull(columns);
                var column = columns.Elements<Column>().First();
                Assert.True(column.BestFit.Value);
                Assert.True(column.Width.HasValue && column.Width.Value > 0);

                var sheetFormat = wsPart.Worksheet.GetFirstChild<SheetFormatProperties>();
                Assert.NotNull(sheetFormat);
                Assert.True(sheetFormat.CustomHeight);
                Assert.True(sheetFormat.DefaultRowHeight > 0);

                var row1 = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex == 1);
                Assert.True(row1.CustomHeight);
                Assert.True(row1.Height.HasValue && row1.Height.Value > 0);

                var row3 = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex == 3);
                Assert.True(row3.CustomHeight);
                Assert.True(row3.Height.HasValue && row3.Height.Value > row1.Height.Value);
            }
        }

        [Fact]
        public void Test_AutoFitRows_EmptyRowsRetainDefaultHeight() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.Empty.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Content");
                sheet.CellValue(2, 1, " ");
                sheet.AutoFitRows();
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                var sheetFormat = wsPart.Worksheet.GetFirstChild<SheetFormatProperties>();
                Assert.NotNull(sheetFormat);
                Assert.True(sheetFormat.CustomHeight);
                Assert.Equal(15.0, sheetFormat.DefaultRowHeight.Value);

                var row2 = wsPart.Worksheet.Descendants<Row>().FirstOrDefault(r => r.RowIndex == 2);
                Assert.NotNull(row2);
                Assert.False(row2!.CustomHeight?.Value ?? false);
                Assert.False(row2.Height?.HasValue ?? false);
            }
        }

        [Fact]
        public void Test_AutoFitRows_RemovesCustomHeightWhenCleared() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.ClearRow.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Content");
                sheet.AutoFitRows();
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets.First();
                sheet.CellValue(1, 1, string.Empty);
                sheet.AutoFitRows();
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                var row1 = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex == 1);
                Assert.False(row1.CustomHeight?.Value ?? false);
                Assert.False(row1.Height?.HasValue ?? false);
            }
        }

        [Fact]
        public void Test_AutoFitColumns_RemovesCustomWidthWhenCleared() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.ClearColumn.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Long text");
                sheet.AutoFitColumns();
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets.First();
                sheet.CellValue(1, 1, string.Empty);
                sheet.AutoFitColumns();
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                var columns = wsPart.Worksheet.GetFirstChild<Columns>();
                Assert.True(columns == null || !columns.Elements<Column>().Any(c => c.Min == 1 && c.Max == 1));
            }
        }

        [Fact]
        public void Test_AutoFitSingleColumn_DoesNotAffectOthers() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.SingleColumn.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Very long text that should expand the column");
                sheet.CellValue(1, 2, "Short");
                sheet.AutoFitColumn(1);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                var columns = wsPart.Worksheet.GetFirstChild<Columns>();
                Assert.NotNull(columns);
                var column1 = columns.Elements<Column>().FirstOrDefault(c => c.Min == 1 && c.Max == 1);
                Assert.NotNull(column1);
                Assert.True(column1!.BestFit?.Value ?? false);
                Assert.True(column1.Width?.Value > 0);

                var column2 = columns.Elements<Column>().FirstOrDefault(c => c.Min == 2 && c.Max == 2);
                Assert.Null(column2);
            }
        }

        [Fact]
        public void Test_AutoFitSingleRow_DoesNotAffectOthers() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.SingleRow.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Line1\nLine2\nLine3");
                sheet.CellValue(2, 1, "Line1\nLine2\nLine3");
                sheet.AutoFitRow(1);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                var row1 = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex == 1);
                Assert.True(row1.CustomHeight?.Value ?? false);
                Assert.True(row1.Height?.Value > 0);

                var row2 = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex == 2);
                Assert.False(row2.CustomHeight?.Value ?? false);
                Assert.False(row2.Height?.HasValue ?? false);
            }
        }

        [Fact]
        public void Test_AutoFit_MixedFonts() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.MixedFonts.xlsx");

            static uint AddFontStyle(SpreadsheetDocument doc, string name, double size) {
                var stylesPart = doc.WorkbookPart.WorkbookStylesPart ?? doc.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                if (stylesPart.Stylesheet == null) {
                    stylesPart.Stylesheet = new Stylesheet(new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font()), new Fills(new Fill()), new Borders(new Border()), new CellFormats(new CellFormat()));
                    stylesPart.Stylesheet.Fonts.Count = 1;
                    stylesPart.Stylesheet.Fills.Count = 1;
                    stylesPart.Stylesheet.Borders.Count = 1;
                    stylesPart.Stylesheet.CellFormats.Count = 1;
                }
                var ss = stylesPart.Stylesheet;
                ss.Fonts.Append(new DocumentFormat.OpenXml.Spreadsheet.Font(new FontName { Val = name }, new FontSize { Val = size }));
                ss.Fonts.Count = (uint)ss.Fonts.ChildElements.Count;
                ss.CellFormats.Append(new CellFormat { FontId = ss.Fonts.Count - 1, ApplyFont = true });
                ss.CellFormats.Count = (uint)ss.CellFormats.ChildElements.Count;
                stylesPart.Stylesheet.Save();
                return ss.CellFormats.Count - 1;
            }

            static void SetCellStyle(SpreadsheetDocument doc, string cellRef, uint styleIndex) {
                var wsPart = doc.WorkbookPart.WorksheetParts.First();
                var cell = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == cellRef);
                cell.StyleIndex = styleIndex;
                wsPart.Worksheet.Save();
            }

            var families = SystemFonts.Collection.Families.ToList();
            string fontName = families.Count > 1 ? families[1].Name : families[0].Name;

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Small");
                sheet.CellValue(2, 1, "Large text");
                sheet.CellValue(3, 1, "Short");
                sheet.CellValue(3, 2, "Tall\nText");

                uint style = AddFontStyle(document._spreadSheetDocument, fontName, 20);
                SetCellStyle(document._spreadSheetDocument, "A2", style);
                SetCellStyle(document._spreadSheetDocument, "B3", style);

                sheet.AutoFitColumn(1);
                sheet.AutoFitRow(3);
                document.Save();
            }

            var font = SystemFonts.CreateFont(fontName, 20);
            var options = new TextOptions(font);
            float zero = TextMeasurer.MeasureSize("0", options).Width;
            double expectedWidth = TextMeasurer.MeasureSize("Large text", options).Width / zero + 1;
            double lineHeight = TextMeasurer.MeasureSize("Tall", options).Height * 72.0 / options.Dpi;
            double expectedHeight = lineHeight * 2 + 2;

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                var column = wsPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().First(c => c.Min == 1 && c.Max == 1);
                Assert.InRange(column.Width!.Value, expectedWidth - 1, expectedWidth + 1);

                var row = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex == 3);
                Assert.InRange(row.Height!.Value, expectedHeight - 1, expectedHeight + 1);
            }
        }

        [Fact]
        public async Task Test_AutoFitConcurrentOperations_AreThreadSafe() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.ConcurrentOperations.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Long piece of text");
                sheet.CellValue(2, 1, "Second line\nwith newline");
                sheet.CellValue(3, 1, "Line1\nLine2\nLine3");

                var tasks = Enumerable.Range(0, 10)
                    .SelectMany(_ => new[] {
                        Task.Run(() => sheet.AutoFitColumns()),
                        Task.Run(() => sheet.AutoFitRows())
                    })
                    .ToArray();
                await Task.WhenAll(tasks);

                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                var columns = wsPart.Worksheet.GetFirstChild<Columns>();
                Assert.NotNull(columns);
                var column = columns.Elements<Column>().First();
                Assert.True(column.BestFit.Value);
                Assert.True(column.Width.HasValue && column.Width.Value > 0);

                var sheetFormat = wsPart.Worksheet.GetFirstChild<SheetFormatProperties>();
                Assert.NotNull(sheetFormat);
                Assert.True(sheetFormat.CustomHeight);
                Assert.True(sheetFormat.DefaultRowHeight > 0);
            }
        }

        [Fact]
        public async Task Test_AutoFitColumnRowConcurrentCalls_AreThreadSafe() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.ConcurrentSingle.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Long piece of text");
                sheet.CellValue(2, 1, "Line1\nLine2");

                var tasks = Enumerable.Range(0, 10)
                    .SelectMany(_ => new[] {
                        Task.Run(() => sheet.AutoFitColumn(1)),
                        Task.Run(() => sheet.AutoFitRow(2))
                    })
                    .ToArray();
                await Task.WhenAll(tasks);

                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                var column = wsPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().First(c => c.Min == 1 && c.Max == 1);
                Assert.True(column.Width.HasValue && column.Width.Value > 0);

                var row = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex == 2);
                Assert.True(row.CustomHeight?.Value ?? false);
                Assert.True(row.Height.HasValue && row.Height.Value > 0);
            }
        }
    }
}
