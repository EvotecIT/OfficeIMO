using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Features;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Layout_setters_write_the_worksheet_part_at_the_document_save_only() {
            string path = Path.Combine(_directoryWithFiles, "LayoutPartSaves.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Report");
                for (int row = 1; row <= 40; row++) {
                    sheet.CellValue(row, 1, "Row " + row);
                    sheet.CellValue(row, 2, row);
                }
                Func<int> saves = TrackPartSaves(document, document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single());

                for (int row = 2; row <= 40; row++) sheet.SetRowHeight(row, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(30)));
                sheet.SetRowHidden(40, true);
                sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(32));
                sheet.SetColumnHidden(3, true);
                sheet.SetColumnWidths(new Dictionary<int, ExcelColumnWidth> { [2] = ExcelColumnWidth.Characters(14) });
                sheet.SetOrientation(OfficePageOrientation.Landscape);
                sheet.SetPaperSize(ExcelPaperSize.A4);
                sheet.SetMarginsPreset(ExcelMarginPreset.Narrow);
                sheet.SetPageSetup(fitToWidth: 1, fitToHeight: 0);
                sheet.SetGridlinesVisible(false);
                sheet.SetRowColumnHeadingsVisible(false);
                sheet.SetZeroValuesVisible(false);
                sheet.SetRightToLeft(true);
                sheet.SetTabColor("FF0000");

                Assert.Equal(0, saves());
                document.Save();
            }

            using SpreadsheetDocument package = SpreadsheetDocument.Open(path, false);
            Worksheet worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            Row[] rows = worksheet.GetFirstChild<SheetData>()!.Elements<Row>().ToArray();
            Assert.Equal(Enumerable.Range(1, 40).Select(index => (uint)index), rows.Select(row => row.RowIndex!.Value));
            Assert.All(rows.Skip(1), row => Assert.Equal(30D, row.Height!.Value));
            Assert.True(rows[39].Hidden!.Value);
            Column[] columns = worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToArray();
            Assert.Equal(32D, columns.Single(column => column.Min!.Value == 1).Width!.Value);
            Assert.Equal(14D, columns.Single(column => column.Min!.Value == 2).Width!.Value);
            Assert.True(columns.Single(column => column.Min!.Value == 3).Hidden!.Value);
            PageSetup pageSetup = worksheet.GetFirstChild<PageSetup>()!;
            Assert.Equal(OrientationValues.Landscape, pageSetup.Orientation!.Value);
            Assert.Equal(1U, pageSetup.FitToWidth!.Value);
            Assert.Equal(0.25D, worksheet.GetFirstChild<PageMargins>()!.Left!.Value);
            SheetView view = worksheet.GetFirstChild<SheetViews>()!.Elements<SheetView>().Single();
            Assert.False(view.ShowGridLines!.Value);
            Assert.False(view.ShowRowColHeaders!.Value);
            Assert.False(view.ShowZeros!.Value);
            Assert.True(view.RightToLeft!.Value);
            Assert.Equal("FFFF0000", worksheet.GetFirstChild<SheetProperties>()!.TabColor!.Rgb!.Value);
        }

        [Fact]
        public void Set_row_height_keeps_rows_in_order_for_any_access_pattern() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet sheet = document.AddWorksheet("Rows");
                foreach (int row in new[] { 2, 5, 9 }) sheet.CellValue(row, 1, "Row " + row);
                foreach (int row in new[] { 10, 3, 7, 1, 9, 4, 12, 6 }) sheet.SetRowHeight(row, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(20 + row)));
                sheet.InsertRows(4, 2);
                foreach (int row in new[] { 4, 15, 8 }) sheet.SetRowHeight(row, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(50 + row)));
                document.Save(stream);
            }

            // Rows 4 and below moved down by two when the rows were inserted; rows 4, 8 and 15 were then set again.
            var expected = new Dictionary<uint, double> {
                [1] = 21, [3] = 23, [4] = 54, [6] = 24, [8] = 58, [9] = 27, [11] = 29, [12] = 30, [14] = 32, [15] = 65
            };
            stream.Position = 0;
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false)) {
                Row[] rows = package.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().ToArray();
                uint[] indexes = rows.Select(row => row.RowIndex!.Value).ToArray();
                Assert.Equal(indexes.Distinct().OrderBy(index => index), indexes);
                Assert.Equal(expected, rows.Where(row => row.Height != null).ToDictionary(row => row.RowIndex!.Value, row => row.Height!.Value));
            }

            stream.Position = 0;
            using var reopened = ExcelDocument.Load(stream);
            Assert.Equal("Row 2", reopened.Sheets[0].CellAt(2, 1).GetValue<string>());
            Assert.Equal("Row 5", reopened.Sheets[0].CellAt(7, 1).GetValue<string>());
            Assert.Equal("Row 9", reopened.Sheets[0].CellAt(11, 1).GetValue<string>());
        }

        [Fact]
        public void Batch_styling_matches_per_call_styling_and_saves_the_stylesheet_once() {
            string[] fills = { "#F8D7DA", "#E6F4EA", "#FFF3CD" };
            void WriteRow(ExcelSheet sheet, int row) {
                for (int column = 1; column <= 3; column++) {
                    sheet.CellValue(row, column, row % 7 == 0 && column == 3 ? "Two" + Environment.NewLine + "lines" : "Value " + row);
                }
                sheet.CellAt(row, 1).SetFillColor(fills[row % fills.Length]).SetBold(true);
                sheet.CellAt(row, 2).SetItalic(true).SetUnderline(true).SetFontColor("#1F4E79");
                sheet.CellAt(row, 3).SetBold(row % 2 == 0);
            }
            (string Worksheet, string Styles, int StyleSaves) Write(bool batch) {
                using var stream = new MemoryStream();
                int styleSaves;
                using (var document = ExcelDocument.Create(stream)) {
                    ExcelSheet sheet = document.AddWorksheet("Styled");
                    for (int column = 1; column <= 3; column++) {
                        sheet.CellValue(1, column, "Header " + column);
                        sheet.CellAt(1, column).HeaderStyle();
                    }
                    Func<int> saves = TrackPartSaves(document, document.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!);
                    if (batch) {
                        sheet.Batch(target => {
                            for (int row = 2; row <= 60; row++) WriteRow(target, row);
                        });
                    } else {
                        for (int row = 2; row <= 60; row++) WriteRow(sheet, row);
                    }
                    for (int column = 1; column <= 3; column++) sheet.WrapCells(1,60,column);
                    styleSaves = saves();
                    document.Save(stream);
                }

                stream.Position = 0;
                using SpreadsheetDocument package = SpreadsheetDocument.Open(stream, false);
                return (package.WorkbookPart!.WorksheetParts.Single().Worksheet!.OuterXml,
                    package.WorkbookPart.WorkbookStylesPart!.Stylesheet!.OuterXml,
                    styleSaves);
            }

            var perCall = Write(batch: false);
            var batched = Write(batch: true);

            Assert.Equal(perCall.Worksheet, batched.Worksheet);
            Assert.Equal(perCall.Styles, batched.Styles);
            Assert.Contains("wrapText=\"1\"", batched.Styles, StringComparison.Ordinal);
            // The batch saves the stylesheet when it ends, and each WrapCells range saves it once.
            Assert.Equal(1 + 3, batched.StyleSaves);
            Assert.True(perCall.StyleSaves > 100, $"Per-call styling saved the stylesheet {perCall.StyleSaves} times.");
        }

        [Fact]
        public void Batch_leaves_an_unstyled_stylesheet_alone_and_range_styling_joins_another_sheets_batch() {
            using var stream = new MemoryStream();
            using (var document = ExcelDocument.Create(stream)) {
                ExcelSheet data = document.AddWorksheet("Data");
                ExcelSheet summary = document.AddWorksheet("Summary");
                summary.CellValue(1, 1, "Total");
                summary.CellValue(1, 2, 5);
                // The workbook gets its styles part with the first styled cell.
                summary.CellAt(1, 1).SetFontColor("#1F4E79");
                Func<int> saves = TrackPartSaves(document, document.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!);

                data.Batch(sheet => {
                    for (int row = 1; row <= 20; row++) sheet.CellValue(row, 1, row);
                });
                Assert.Equal(0, saves());

                data.Batch(_ => summary.Range("A1:B1").SetBold());
                Assert.Equal(1, saves());
                document.Save(stream);
            }

            stream.Position = 0;
            using var reopened = ExcelDocument.Load(stream);
            Assert.Equal(20D, reopened.Sheets[0].CellAt(20, 1).GetValue<double>());
            Assert.True(reopened.Sheets[1].CellAt(1, 1).GetStyle().Bold);
            Assert.True(reopened.Sheets[1].CellAt(1, 2).GetStyle().Bold);
        }

        private static Func<int> TrackPartSaves(ExcelDocument document, OpenXmlPart part) {
            Assert.NotNull(part);
            SpreadsheetDocument package = document.OpenXmlDocument;
            if (package.Features.Get<IPartRootEventsFeature>() == null) package.AddPartRootEventsFeature();
            int saves = 0;
            package.Features.Get<IPartRootEventsFeature>()!.Change += args => {
                if (args.Type == EventType.Saved && ReferenceEquals(args.Argument, part)) saves++;
            };
            return () => saves;
        }
    }
}
