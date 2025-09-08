using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests for setting cell values in Excel sheets.
    /// </summary>
    public partial class Excel {
        [Fact]
        public void Test_CellValues() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValues.xlsx");
            var date = new DateTime(2024, 1, 2);
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Hello");
                sheet.CellValue(2, 1, 123.45);
                sheet.CellValue(3, 1, 678.90m);
                sheet.CellValue(4, 1, date);
                sheet.CellValue(5, 1, true);
                sheet.CellFormula(6, 1, "SUM(A2:A3)");
                sheet.Cell(7, 1, 1.23, "A2+1", "0.00");
                document.Save();
            }

            SpreadsheetDocument spreadsheet = null!;
            Exception? ex = Record.Exception(() => spreadsheet = SpreadsheetDocument.Open(filePath, false));
            Assert.Null(ex);
            using (spreadsheet) {
                  ValidateSpreadsheetDocument(filePath, spreadsheet);

                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cells = wsPart.Worksheet.Descendants<Cell>().ToList();
                SharedStringTablePart shared = spreadsheet.WorkbookPart!.SharedStringTablePart!;
                Assert.NotNull(shared);
                Assert.Equal("Hello", shared.SharedStringTable!.ElementAt(0).InnerText);

                Cell cellString = cells.First(c => c.CellReference == "A1");
                  Assert.Equal(CellValues.SharedString, cellString.DataType!.Value);
                  Assert.Equal("0", cellString.CellValue!.Text);

                Cell cellDouble = cells.First(c => c.CellReference == "A2");
                  Assert.Equal(CellValues.Number, cellDouble.DataType!.Value);
                  Assert.Equal("123.45", cellDouble.CellValue!.Text);

                Cell cellDecimal = cells.First(c => c.CellReference == "A3");
                  Assert.Equal(CellValues.Number, cellDecimal.DataType!.Value);
                  Assert.Equal("678.90", cellDecimal.CellValue!.Text);

                Cell cellDate = cells.First(c => c.CellReference == "A4");
                  Assert.Equal(CellValues.Number, cellDate.DataType!.Value);
                  Assert.Equal(date.ToOADate().ToString(CultureInfo.InvariantCulture), cellDate.CellValue!.Text);

                Cell cellBool = cells.First(c => c.CellReference == "A5");
                  Assert.Equal(CellValues.Boolean, cellBool.DataType!.Value);
                  Assert.Equal("1", cellBool.CellValue!.Text);

                Cell cellFormula = cells.First(c => c.CellReference == "A6");
                  Assert.NotNull(cellFormula.CellFormula);
                  Assert.Equal("SUM(A2:A3)", cellFormula.CellFormula!.Text);

                Cell cellCombined = cells.First(c => c.CellReference == "A7");
                  Assert.Equal(CellValues.Number, cellCombined.DataType!.Value);
                  Assert.Equal("1.23", cellCombined.CellValue!.Text);
                  Assert.NotNull(cellCombined.CellFormula);
                  Assert.Equal("A2+1", cellCombined.CellFormula!.Text);
                  Assert.NotNull(cellCombined.StyleIndex);
                var styles = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
                var formats = styles.NumberingFormats!.Elements<NumberingFormat>().ToList();
                Assert.Contains(formats, n => n.FormatCode != null && n.FormatCode.Value == "0.00");
            }
        }
    }
}
