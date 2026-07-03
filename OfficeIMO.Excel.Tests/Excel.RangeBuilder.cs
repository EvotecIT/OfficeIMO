using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelRangeBuilderTests {
        private static string GetCellValue(SpreadsheetDocument document, WorksheetPart worksheetPart, string cellReference) {
            var cell = worksheetPart.Worksheet.Descendants<Cell>().First(c => c.CellReference != null && c.CellReference.Value == cellReference);
            var value = cell.CellValue?.Text ?? string.Empty;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
                var table = document.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
                if (table != null && int.TryParse(value, out int id)) {
                    return table.ChildElements[id].InnerText;
                }
            }
            return value;
        }

        [Fact]
        public void RangeBuilderSetsValues() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            object[,] values = { { "A", "B" }, { "C", "D" } };

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AsFluent().Sheet("Data", s => s.Range("A1:B2", r => r.Set(values)));
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                Assert.NotNull(document._spreadSheetDocument);
                var workbookPart = document._spreadSheetDocument.WorkbookPart;
                Assert.NotNull(workbookPart);
                var sheetPart = workbookPart.WorksheetParts.First();
                Assert.Equal("A", GetCellValue(document._spreadSheetDocument, sheetPart, "A1"));
                Assert.Equal("D", GetCellValue(document._spreadSheetDocument, sheetPart, "B2"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void RangeBuilderSetsSingleCell() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AsFluent().Sheet("Data", s => s.Range("A1:C3", r => r.Cell(2, 2, "X")));
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                Assert.NotNull(document._spreadSheetDocument);
                var workbookPart = document._spreadSheetDocument.WorkbookPart;
                Assert.NotNull(workbookPart);
                var sheetPart = workbookPart.WorksheetParts.First();
                Assert.Equal("X", GetCellValue(document._spreadSheetDocument, sheetPart, "B2"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void RangeBuilderAppliesNumberFormat() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            object[,] values = { { 1.2, 3.4 }, { 5.6, 7.8 } };

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AsFluent().Sheet("Data", s => s.Range("A1:B2", r => {
                    r.Set(values);
                    r.NumberFormat("0.00");
                }));
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = spreadsheet.WorkbookPart;
                Assert.NotNull(workbookPart);
                var wsPart = workbookPart.WorksheetParts.First();
                var cell = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");
                Assert.NotNull(cell.StyleIndex);
                uint styleIndex = cell.StyleIndex.Value;
                var stylesPart = workbookPart.WorkbookStylesPart;
                Assert.NotNull(stylesPart);
                var styles = stylesPart.Stylesheet;
                Assert.NotNull(styles);
                Assert.NotNull(styles.CellFormats);
                var cellFormat = (CellFormat)styles.CellFormats.ElementAt((int)styleIndex);
                Assert.NotNull(cellFormat.NumberFormatId);
                var nfId = cellFormat.NumberFormatId.Value;
                var numberingFormats = styles.NumberingFormats;
                Assert.NotNull(numberingFormats);
                var numberingFormat = numberingFormats.Elements<NumberingFormat>().First(n => n.NumberFormatId != null && n.NumberFormatId.Value == nfId);
                Assert.NotNull(numberingFormat.FormatCode);
                Assert.Equal("0.00", numberingFormat.FormatCode.Value);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void RangeBuilderCanClearValues() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            object[,] values = { { "A", "B" }, { "C", "D" } };

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AsFluent().Sheet("Data", s => s.Range("A1:B2", r => {
                    r.Set(values);
                    r.Clear();
                }));
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                Assert.NotNull(document._spreadSheetDocument);
                var workbookPart = document._spreadSheetDocument.WorkbookPart;
                Assert.NotNull(workbookPart);
                var sheetPart = workbookPart.WorksheetParts.First();
                Assert.Equal(string.Empty, GetCellValue(document._spreadSheetDocument, sheetPart, "A1"));
                Assert.Equal(string.Empty, GetCellValue(document._spreadSheetDocument, sheetPart, "B2"));
            }

            File.Delete(filePath);
        }
    }
}

