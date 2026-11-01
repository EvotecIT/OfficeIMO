using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        public static IEnumerable<object[]> BufferedCellShapes() {
            yield return new object[] { "trailing total", (Action<ExcelSheet>)(sheet => {
                WriteBufferedRows(sheet, firstRow: 1, rows: 100);
                sheet.CellValue(101, 1, "Total");
                sheet.CellFormula(101, 2, "SUM(B1:B100)");
            }) };
            yield return new object[] { "title above data", (Action<ExcelSheet>)(sheet => {
                sheet.CellValue(2, 2, "Title");
                WriteBufferedRows(sheet, firstRow: 4, rows: 100);
            }) };
            yield return new object[] { "rows and columns with gaps", (Action<ExcelSheet>)(sheet => {
                for (int row = 2; row <= 400; row += 3) {
                    for (int column = 2; column <= 12; column += 4) {
                        sheet.CellValue(row, column, row * column);
                    }
                }
            }) };
            yield return new object[] { "explicit blanks", (Action<ExcelSheet>)(sheet => {
                for (int row = 1; row <= 70; row++) {
                    sheet.CellValue(row, 1, (object?)null);
                    sheet.CellValue(row, 3, "kept");
                }
            }) };
        }

        private static void WriteBufferedRows(ExcelSheet sheet, int firstRow, int rows) {
            for (int index = 0; index < rows; index++) {
                sheet.CellValue(firstRow + index, 1, "name " + index);
                sheet.CellValue(firstRow + index, 2, index * 1.5);
                sheet.CellValue(firstRow + index, 3, new DateTime(2024, 1, 1).AddDays(index));
            }
        }

        [Theory]
        [MemberData(nameof(BufferedCellShapes))]
        public void CellValueBuffer_ShapesSaveDirectlyWithTheContentOfTheDomRoute(string shape, Action<ExcelSheet> write) {
            using MemoryStream direct = SaveCells(write, options: null, ExcelSavePackageWriter.DirectDataSetPackage);
            using MemoryStream dom = SaveCells(write, new ExcelSaveOptions { DisableFastPackageWriter = true }, expectedWriter: null);

            using ExcelDocument reopened = ExcelDocument.Load(direct);
            using ExcelDocument expected = ExcelDocument.Load(dom);
            Assert.Empty(reopened.ValidateOpenXml());
            Assert.True(ReadCells(expected).Count > 0, shape);
            Assert.Equal(ReadCells(expected), ReadCells(reopened));
        }

        private static MemoryStream SaveCells(Action<ExcelSheet> write, ExcelSaveOptions? options, ExcelSavePackageWriter? expectedWriter) {
            using ExcelDocument document = ExcelDocument.Create();
            write(document.AddSheet("Data"));
            var output = new MemoryStream();
            document.Save(output, options);
            if (expectedWriter != null) {
                Assert.Equal(expectedWriter, document.LastSaveDiagnostics.Writer);
            }

            output.Position = 0;
            return output;
        }

        private static List<string> ReadCells(ExcelDocument document) {
            ExcelSheet sheet = document.Sheets[0];
            return sheet.EnumerateCells()
                .Select(cell => $"{cell.Row}:{cell.Column}:{cell.Value}:{sheet.GetFormulaText(cell.Row, cell.Column)}")
                .ToList();
        }

        [Fact]
        public void CellValueBuffer_FarSparseCellsLeaveTheRectangleWriters() {
            Action<ExcelSheet> write = sheet => {
                WriteBufferedRows(sheet, firstRow: 1, rows: 100);
                sheet.CellValue(ExcelA1.MaxRows, ExcelA1.MaxColumns, "corner");
            };

            using MemoryStream output = SaveCells(write, options: null, expectedWriter: null);
            using ExcelDocument reopened = ExcelDocument.Load(output);

            var corner = reopened.Sheets[0].EnumerateCells().Last();
            Assert.Equal((ExcelA1.MaxRows, ExcelA1.MaxColumns, "corner"), (corner.Row, corner.Column, corner.Value));
        }

        [Fact]
        public void CellValueBuffer_MergesAndCommentsSurviveTheMoveToTheDom() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            WriteBufferedRows(sheet, firstRow: 1, rows: 100);
            sheet.MergeRange("E1:F2");
            sheet.SetComment(1, 1, "note");

            using var output = new MemoryStream();
            document.Save(output);
            output.Position = 0;

            using ExcelDocument reopened = ExcelDocument.Load(output);
            ExcelSheet saved = reopened.Sheets[0];
            Assert.Single(saved.GetMergedRanges());
            Assert.Single(saved.GetComments());
            var worksheet = reopened.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            Cell date = worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "C1");
            var styles = reopened.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            CellFormat format = styles.CellFormats!.Elements<CellFormat>().ElementAt((int)date.StyleIndex!.Value);
            Assert.Equal(14U, format.NumberFormatId!.Value);
        }
    }
}
