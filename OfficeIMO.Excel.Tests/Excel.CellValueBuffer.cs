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
                sheet.Cell(101, 1).SetValue("Total");
                sheet.Cell(101, 2).SetFormula("SUM(B1:B100)");
            }) };
            yield return new object[] { "title above data", (Action<ExcelSheet>)(sheet => {
                sheet.Cell(2, 2).SetValue("Title");
                WriteBufferedRows(sheet, firstRow: 4, rows: 100);
            }) };
            yield return new object[] { "rows and columns with gaps", (Action<ExcelSheet>)(sheet => {
                for (int row = 2; row <= 400; row += 3) {
                    for (int column = 2; column <= 12; column += 4) {
                        sheet.Cell(row, column).SetValue(row * column);
                    }
                }
            }) };
            yield return new object[] { "explicit blanks", (Action<ExcelSheet>)(sheet => {
                for (int row = 1; row <= 70; row++) {
                    sheet.Cell(row, 1).SetValue((object?)null);
                    sheet.Cell(row, 3).SetValue("kept");
                }
            }) };
        }

        private static void WriteBufferedRows(ExcelSheet sheet, int firstRow, int rows) {
            for (int index = 0; index < rows; index++) {
                sheet.Cell(firstRow + index, 1).SetValue("name " + index);
                sheet.Cell(firstRow + index, 2).SetValue(index * 1.5);
                sheet.Cell(firstRow + index, 3).SetValue(new DateTime(2024, 1, 1).AddDays(index));
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
            summary.Cell(1, 1).SetValue("Total");
            summary.Cell(1, 2).SetFormula("SUM(Data!B1:B100)");
            WriteBufferedRows(data, firstRow: 1, rows: 100);
            data.Cell(101, 1).SetValue("line one\nline two");
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
                styled.Cell(1, 1).SetValue("Header");
                styled.Cell(1, 1).SetBold();
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
                sheet.Range("E1:F2").Merge();
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
        public void CellValueBuffer_LineBreakKeepsWrappedTextOutOfTheDom() {
            using MemoryStream output = SaveCells(
                sheet => {
                    WriteBufferedRows(sheet, firstRow: 1, rows: 100);
                    sheet.Cell(101, 1).SetValue("line one\nline two");
                },
                options: null,
                ExcelSavePackageWriter.ExtendedPackage);

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
                    sheet.Cell(row, 1).SetValue("item number " + row + (row % 97 == 0 ? " WWWWWWWW" : string.Empty));
                    sheet.Cell(row, 2).SetValue(row * 1.25);
                    sheet.Cell(row, 3).SetValue(new DateTime(2024, 1, 1).AddDays(row));
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

        private static CellFormat FormatOf(ExcelDocument document, string reference) {
            var workbookPart = document.OpenXmlDocument.WorkbookPart!;
            Cell cell = document.Sheets[0].WorksheetPart.Worksheet!.Descendants<Cell>()
                .Single(candidate => candidate.CellReference?.Value == reference);
            return workbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>()
                .ElementAt((int)cell.StyleIndex!.Value);
        }

        [Fact]
        public void CellValueBuffer_CellFormatsStayWithBufferedCells() {
            using MemoryStream output = SaveWorkbook(document => {
                ExcelSheet styled = document.AddSheet("Styled");
                ExcelSheet plain = document.AddSheet("Plain");
                styled.Cell(1, 1).SetValue("Header");
                styled.Cell(1, 1).SetBold();
                WriteBufferedRows(styled, firstRow: 2, rows: 200);
                styled.Cell(150, 2).SetNumberFormat("0.000");
                styled.Cell(150, 3).SetFillColor("#DDEBF7");
                styled.Cell(160, 1).SetBold();
                styled.Cell(160, 1).SetValue("overwritten");
                WriteBufferedRows(plain, firstRow: 1, rows: 100);
                Assert.True(document.HasPendingDirectCellValues);
            }, options: null);

            using ExcelDocument reopened = ExcelDocument.Load(output);
            Assert.Empty(reopened.ValidateOpenXml());
            var styles = reopened.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            Assert.NotNull(styles.Fonts!.Elements<Font>().ElementAt((int)FormatOf(reopened, "A1").FontId!.Value).Bold);
            Assert.NotNull(styles.Fonts!.Elements<Font>().ElementAt((int)FormatOf(reopened, "A160").FontId!.Value).Bold);
            CellFormat dated = FormatOf(reopened, "C150");
            Assert.Equal(14U, dated.NumberFormatId!.Value);
            Assert.NotEqual(0U, dated.FillId!.Value);
            Assert.Contains(reopened.Sheets[0].EnumerateCells(), cell => cell.Row == 160 && Equals(cell.Value, "overwritten"));
        }

        [Fact]
        public void CellValueBuffer_LargeSheetKeepsAFormatOnItsLastCells() {
            const int rows = 15000;
            using MemoryStream output = SaveWorkbook(document => {
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Header");
                sheet.Cell(1, 1).SetBold();
                WriteBufferedRows(sheet, firstRow: 2, rows);
                sheet.Cell(rows + 1, 2).SetNumberFormat("0.000");
            }, options: null);

            using ExcelDocument reopened = ExcelDocument.Load(output);
            Assert.Equal((uint)reopened.Sheets[0].EnumerateCells().Count(), (uint)(rows * 3 + 1));
            Assert.Equal("0.000", reopened.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.NumberingFormats!
                .Elements<NumberingFormat>().Single(format => format.NumberFormatId!.Value == FormatOf(reopened, "B" + (rows + 1)).NumberFormatId!.Value).FormatCode!.Value);
        }

        [Fact]
        public void CellValueBuffer_HoldsMoreFormatsThanTwelveBitsCanIndex() {
            const int formats = 4200;
            using MemoryStream output = SaveWorkbook(document => {
                ExcelSheet sheet = document.AddSheet("Data");
                for (int row = 1; row <= formats; row++) {
                    sheet.Cell(row, 1).SetValue(row);
                    sheet.Cell(row, 1).SetNumberFormat("0\"" + row + "\"");
                }

                Assert.True(document.HasPendingDirectCellValues);
            }, options: null);

            using ExcelDocument reopened = ExcelDocument.Load(output);
            var stylesheet = reopened.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            var cells = reopened.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>().ToArray();
            CellFormat[] cellFormats = stylesheet.CellFormats!.Elements<CellFormat>().ToArray();
            var codes = stylesheet.NumberingFormats!.Elements<NumberingFormat>().ToDictionary(format => format.NumberFormatId!.Value, format => format.FormatCode!.Value);
            Assert.Equal(formats, cells.Length);
            Assert.Equal(formats, cells.Select(cell => cell.StyleIndex!.Value).Distinct().Count());
            Assert.Equal("0\"" + formats + "\"", codes[cellFormats[(int)cells[formats - 1].StyleIndex!.Value].NumberFormatId!.Value]);
        }

        [Fact]
        public void CellValueBuffer_RowSettingsTablesRulesChartsAndPropertiesStayOutOfTheDom() {
            using MemoryStream output = SaveWorkbook(document => {
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.Cell(1, 1).SetValue("Name");
                sheet.Cell(1, 2).SetValue("Amount");
                sheet.Cell(1, 3).SetValue("When");
                WriteBufferedRows(sheet, firstRow: 2, rows: 200);
                sheet.SetRowHeight(150, ExcelColumnWidth.Characters(30));
                sheet.SetRowHidden(151, true);
                sheet.AddTable("A1:C201", true, "Sales", ExcelTableStyle.TableStyleMedium2);
                sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
                    Range = "B2:B201",
                    Type = "expression",
                    Formulas = new[] { "B2>10" },
                    DifferentialFillColorArgb = "C6EFCE"
                });
                sheet.AddChart(OfficeChartType.Line, new ExcelChartDataRange("Data", 1, 2, 200, 1, true), new ExcelChartOptions { Row = 1, Column = 6 });
                document.BuiltinDocumentProperties.Title = "Sales report";
                Assert.True(document.HasPendingDirectCellValues);
            }, options: null);

            using ExcelDocument reopened = ExcelDocument.Load(output);
            Assert.Empty(reopened.ValidateOpenXml());
            Assert.Equal("Sales report", reopened.BuiltinDocumentProperties.Title);
            var part = reopened.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single();
            Assert.Equal("Sales", part.TableDefinitionParts.Single().Table!.Name!.Value);
            Assert.Single(part.DrawingsPart!.ChartParts);
            Row[] rows = part.Worksheet!.Descendants<Row>().ToArray();
            Assert.True(rows[149].CustomHeight!.Value);
            Assert.True(rows[150].Hidden!.Value);
            ConditionalFormattingRule rule = part.Worksheet.Descendants<ConditionalFormattingRule>().Single();
            Assert.NotNull(reopened.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.DifferentialFormats!
                .Elements<DifferentialFormat>().ElementAt((int)rule.FormatId!.Value));
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
                sheet.Cell(ExcelA1.MaxRows, ExcelA1.MaxColumns).SetValue("corner");
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
            sheet.Range("E1:F2").Merge();
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

        private static void WriteMixedBufferedCells(ExcelSheet sheet) {
            for (int row = 1; row <= 80; row++) {
                sheet.Cell(row, 1).SetValue("text " + row);
                sheet.Cell(row, 2).SetValue(row * 1.25);
                sheet.Cell(row, 3).SetValue(row);
                sheet.Cell(row, 4).SetValue(9_007_199_254_740_993L + row);
                sheet.Cell(row, 5).SetValue(row % 2 == 0);
                sheet.Cell(row, 6).SetValue(new DateTime(2024, 1, 1, 6, 30, 0).AddDays(row));
            }

            sheet.Cell(81, 1).SetValue("  padded  ");
            sheet.Cell(81, 2).SetValue("   ");
            sheet.Cell(81, 3).SetValue("line\nbreak");
            sheet.Cell(81, 4).SetValue(1.5m);
            sheet.Cell(81, 5).SetValue(TimeSpan.FromHours(30));
            sheet.Cell(81, 6).SetValue((object?)null);
            sheet.Cell(81, 7).SetFormula("SUM(B1:B80)");
            sheet.Cell(82, 1).SetValue(3d);
            sheet.Cell(82, 1).SetNumberFormat("yyyy-mm-dd");
        }

        [Fact]
        public void CellValueBuffer_FormulaOverAWrittenValueSavesThatValueAsItsTypedCache() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            for (int row = 1; row <= 200; row++) {
                sheet.Cell(row, 1).SetValue(row * 2d);
                sheet.Cell(row, 2).SetValue(row % 2 == 0);
            }
            // Row 1 is among the cells also written to the DOM; row 150 lives only in the buffer.
            foreach (int row in new[] { 1, 150 }) {
                sheet.Cell(row, 1).SetFormula("A2*0");
                sheet.Cell(row, 2).SetFormula("NOT(TRUE)");
            }

            using var output = new MemoryStream();
            document.Save(output);
            output.Position = 0;
            using ExcelDocument reopened = ExcelDocument.Load(output);
            ExcelSheet saved = reopened.Sheets[0];
            foreach (int row in new[] { 1, 150 }) {
                Assert.Equal("A2*0", saved.GetFormulaText(row, 1));
                Assert.Equal(row * 2d, saved.Cell(row, 1).GetValue<double>());
                Assert.Equal("NOT(TRUE)", saved.GetFormulaText(row, 2));
                Assert.Equal(row % 2 == 0, saved.Cell(row, 2).GetValue<bool>());
            }
        }

        [Fact]
        public void CellValueBuffer_ReadsMatchTheDomRouteAndLeaveTheCellsBuffered() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet buffered = document.AddSheet("Buffered");
            WriteMixedBufferedCells(buffered);
            using ExcelDocument domDocument = ExcelDocument.Create();
            ExcelSheet dom = domDocument.AddSheet("Dom");
            WriteMixedBufferedCells(dom);
            dom.MaterializePendingDirectCellValues();

            Assert.Equal(
                dom.EnumerateCells().Select(cell => (cell.Row, cell.Column, cell.Value)).ToArray(),
                buffered.EnumerateCells().Select(cell => (cell.Row, cell.Column, cell.Value)).ToArray());
            Assert.Equal(
                dom.ToDataTable(headersInFirstRow: false).Rows.Cast<System.Data.DataRow>().Select(row => row.ItemArray).ToArray(),
                buffered.ToDataTable(headersInFirstRow: false).Rows.Cast<System.Data.DataRow>().Select(row => row.ItemArray).ToArray());
            Assert.True(buffered.PendingCellValueCount > 0);
        }

        [Fact]
        public void CellValueBuffer_OpenEnumeratorKeepsTheCellsOfItsRead() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            WriteBufferedRows(sheet, firstRow: 1, rows: 100);

            using (var cells = sheet.EnumerateCells().GetEnumerator()) {
                Assert.True(cells.MoveNext());
                sheet.Cell(100, 1).SetValue("changed");
                sheet.Cell(101, 1).SetValue("added");
                int count = 1;
                object? lastName = null;
                while (cells.MoveNext()) {
                    count++;
                    if (cells.Current.Column == 1) lastName = cells.Current.Value;
                }

                Assert.Equal(300, count);
                Assert.Equal("name 99", lastName);
            }

            Assert.Equal(301, sheet.EnumerateCells().Count());
            Assert.Equal("changed", sheet.EnumerateCells().First(cell => cell.Row == 100).Value);
        }

        [Fact]
        public void CellValueBuffer_ReadAfterEveryWriteMovesToTheDomAndKeepsEveryCell() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            WriteBufferedRows(sheet, firstRow: 1, rows: 30_000);

            for (int read = 1; read <= 20; read++) {
                sheet.Cell(30_000 + read, 1).SetValue(read);
                Assert.Equal(90_000 + read, sheet.EnumerateCells().Count());
            }

            Assert.Equal(0, sheet.PendingCellValueCount);
        }
    }
}
