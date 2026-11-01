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
            return document.Sheets.SelectMany(sheet => sheet.EnumerateCells()
                .Select(cell => $"{sheet.Name}:{cell.Row}:{cell.Column}:{cell.Value}:{sheet.GetFormulaText(cell.Row, cell.Column)}"))
                .ToList();
        }

        private static Action<ExcelDocument> SeveralBufferedSheets => document => {
            ExcelSheet summary = document.AddSheet("Summary");
            ExcelSheet data = document.AddSheet("Data");
            document.AddSheet("Empty");
            summary.CellValue(1, 1, "Total");
            summary.CellFormula(1, 2, "SUM(Data!B1:B100)");
            WriteBufferedRows(data, firstRow: 1, rows: 100);
            data.CellValue(101, 1, "line one\nline two");
            Assert.True(document.HasPendingDirectCellValues);
        };

        private static MemoryStream SaveWorkbook(Action<ExcelDocument> write, ExcelSaveOptions? options) {
            using ExcelDocument document = ExcelDocument.Create();
            write(document);
            var output = new MemoryStream();
            document.Save(output, options);
            output.Position = 0;
            return output;
        }

        [Fact]
        public void CellValueBuffer_SeveralSheetsKeepTheContentOfTheDomRoute() {
            using MemoryStream direct = SaveWorkbook(SeveralBufferedSheets, options: null);
            using MemoryStream dom = SaveWorkbook(SeveralBufferedSheets, new ExcelSaveOptions { DisableFastPackageWriter = true });

            using ExcelDocument reopened = ExcelDocument.Load(direct);
            using ExcelDocument expected = ExcelDocument.Load(dom);
            Assert.Empty(reopened.ValidateOpenXml());
            Assert.Equal(new[] { "Summary", "Data", "Empty" }, reopened.Sheets.Select(sheet => sheet.Name));
            Assert.Equal(ReadCells(expected), ReadCells(reopened));
        }

        [Fact]
        public void CellValueBuffer_StyledSiblingSheetKeepsItsFormats() {
            using MemoryStream output = SaveWorkbook(document => {
                ExcelSheet plain = document.AddSheet("Plain");
                ExcelSheet styled = document.AddSheet("Styled");
                WriteBufferedRows(plain, firstRow: 1, rows: 100);
                styled.CellValue(1, 1, "Header");
                styled.CellBold(1, 1);
            }, options: null);

            using ExcelDocument reopened = ExcelDocument.Load(output);
            var stylesheet = reopened.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            Assert.Contains(stylesheet.Fonts!.Elements<Font>(), font => font.Bold != null);
            Assert.Empty(reopened.ValidateOpenXml());
        }

        [Fact]
        public void CellValueBuffer_WorksheetSettingsLeaveCellsBuffered() {
            using MemoryStream output = SaveWorkbook(document => {
                ExcelSheet sheet = document.AddSheet("Data");
                WriteBufferedRows(sheet, firstRow: 1, rows: 100);
                sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(30));
                sheet.MergeRange("E1:F2");
                sheet.SetComment(1, 1, "note");
                sheet.SetHyperlink(90, 1, "https://example.org", "Home");
                Assert.True(document.HasPendingDirectCellValues);
            }, options: null);

            using ExcelDocument reopened = ExcelDocument.Load(output);
            Assert.Empty(reopened.ValidateOpenXml());
            ExcelSheet saved = reopened.Sheets[0];
            Assert.Single(saved.GetMergedRanges());
            Assert.Single(saved.GetComments());
            var part = reopened.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single();
            Assert.Equal("https://example.org", part.HyperlinkRelationships.Single().Uri.OriginalString);
            Assert.Equal(30D, part.Worksheet!.Descendants<Column>().Single().Width!.Value);
            Cell link = part.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A90");
            var styles = reopened.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            uint fontId = styles.CellFormats!.Elements<CellFormat>().ElementAt((int)link.StyleIndex!.Value).FontId!.Value;
            Assert.Equal("FF0563C1", styles.Fonts!.Elements<Font>().ElementAt((int)fontId).Color!.Rgb!.Value);
        }

        [Fact]
        public void CellValueBuffer_LineBreakKeepsWrappedTextOnTheDirectWriter() {
            using MemoryStream output = SaveCells(
                sheet => {
                    WriteBufferedRows(sheet, firstRow: 1, rows: 100);
                    sheet.CellValue(101, 1, "line one\nline two");
                },
                options: null,
                ExcelSavePackageWriter.DirectDataSetPackage);

            using ExcelDocument reopened = ExcelDocument.Load(output);
            Assert.Empty(reopened.ValidateOpenXml());
            Cell text = reopened.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "A101");
            CellFormat format = reopened.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.CellFormats!
                .Elements<CellFormat>().ElementAt((int)text.StyleIndex!.Value);
            Assert.True(format.Alignment!.WrapText!.Value);
        }

        [Fact]
        public void CellValueBuffer_AutoFitMatchesTheMeasuredDomWidths() {
            static double[] Widths(bool materializeFirst) {
                using ExcelDocument document = ExcelDocument.Create();
                ExcelSheet sheet = document.AddSheet("Data");
                for (int row = 1; row <= 5000; row++) {
                    sheet.CellValue(row, 1, "item number " + row + (row % 97 == 0 ? " WWWWWWWW" : string.Empty));
                    sheet.CellValue(row, 2, row * 1.25);
                    sheet.CellValue(row, 3, new DateTime(2024, 1, 1).AddDays(row));
                }

                if (materializeFirst) {
                    sheet.MaterializePendingDirectCellValues();
                }

                sheet.AutoFitColumns();
                Assert.Equal(!materializeFirst, document.HasPendingDirectCellValues);
                using var output = new MemoryStream();
                document.Save(output);
                output.Position = 0;
                using ExcelDocument reopened = ExcelDocument.Load(output);
                return reopened.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Column>()
                    .Select(column => column.Width!.Value)
                    .ToArray();
            }

            Assert.Equal(Widths(materializeFirst: true), Widths(materializeFirst: false));
        }

        [Fact]
        public void CellValueBuffer_FreezeAndFilterStayOnTheDirectWriter() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            WriteBufferedRows(sheet, firstRow: 1, rows: 100);
            sheet.Freeze(topRows: 1);
            sheet.AutoFilterAdd("A1:C100");

            using var output = new MemoryStream();
            document.Save(output);
            Assert.Equal(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
            output.Position = 0;

            using ExcelDocument reopened = ExcelDocument.Load(output);
            var worksheet = reopened.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            Assert.Equal(1D, worksheet.Descendants<Pane>().Single().VerticalSplit!.Value);
            Assert.Equal("A1:C100", worksheet.GetFirstChild<AutoFilter>()!.Reference!.Value);
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
