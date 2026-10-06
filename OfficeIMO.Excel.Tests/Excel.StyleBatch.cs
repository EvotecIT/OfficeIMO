using System;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml.Features;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Wrap_text_batch_can_run_in_another_sheets_workbook_batch() {
            string path = Path.Combine(_directoryWithFiles, "WrapTextBatch.CrossSheet.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var first = document.AddWorksheet("First");
                var second = document.AddWorksheet("Second");
                first.CellValue(1, 1, "First value");
                second.CellValue(1, 1, "Second value");
                first.Batch(_ => second.CellWrapTextFor(new[] { (1, 1) }));
                document.Save();
            }
            using var reopened = ExcelDocument.Load(path);
            Assert.Equal("First value", reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
            Assert.False(reopened.Sheets[0].CellAt(1, 1).GetStyle().WrapText);
            Assert.Equal("Second value", reopened.Sheets[1].CellAt(1, 1).GetValue<string>());
            Assert.True(reopened.Sheets[1].CellAt(1, 1).GetStyle().WrapText);
        }

        [Fact]
        public void Wrap_text_batch_saves_styles_once_and_preserves_sparse_cell_formatting() {
            string path = Path.Combine(_directoryWithFiles, "WrapTextBatch.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Styled");
                var selected = Enumerable.Range(1, 64).Select(row => (Row: row, Column: row % 2 == 0 ? 3 : 1)).ToArray();
                foreach (var cell in selected) {
                    sheet.CellValue(cell.Row, cell.Column, cell.Row);
                    sheet.CellAt(cell.Row, cell.Column).SetBold();
                    sheet.FormatCell(cell.Row, cell.Column, "0.00");
                }
                var package = document.OpenXmlDocument;
                var stylesPart = package.WorkbookPart!.WorkbookStylesPart!;
                package.AddPartRootEventsFeature();
                int saves = 0;
                package.Features.Get<IPartRootEventsFeature>()!.Change += args => {
                    if (args.Type == EventType.Saved && ReferenceEquals(args.Argument, stylesPart)) saves++;
                };
                sheet.CellWrapTextFor(selected);
                Assert.Equal(1, saves);
                document.Save();
            }
            using var reopened = ExcelDocument.Load(path);
            ExcelSheet saved = reopened.Sheets[0];
            Assert.Equal(64, saved.EnumerateCells().Count());
            foreach (var cell in saved.EnumerateCells()) {
                var style = saved.CellAt(cell.Row, cell.Column).GetStyle();
                Assert.True(style.WrapText && style.Bold);
                Assert.Equal("0.00", style.NumberFormatCode);
            }
        }

        [Fact]
        public void Wrap_text_batch_validates_before_editing_and_can_clear_wrapping() {
            string path = Path.Combine(_directoryWithFiles, "WrapTextBatch.Validation.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Styled");
                sheet.CellValue(1, 1, "Preserved");
                sheet.CellWrapText(1, 1);
                Assert.Throws<ArgumentOutOfRangeException>(() => sheet.CellWrapTextFor(new[] { (1, 1), (0, 2) }, false));
                Assert.True(sheet.CellAt(1, 1).GetStyle().WrapText);
                Assert.Single(sheet.EnumerateCells());
                sheet.Batch(selected => selected.CellWrapTextFor(new[] { (1, 1), (1, 1) }, false));
                document.Save();
            }
            using var reopened = ExcelDocument.Load(path);
            Assert.Equal("Preserved", reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
            Assert.False(reopened.Sheets[0].CellAt(1, 1).GetStyle().WrapText);
        }

        [Fact]
        public void Conversion_style_batch_reuses_distinct_styles_and_saves_valid_workbook() {
            string path = Path.Combine(_directoryWithFiles, "ConversionStyleBatch.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet sheet = document.AddWorksheet("Styled");
                ExcelSheet second = document.AddWorksheet("Second");
                using (sheet.BeginStyleBatch()) {
                    for (int row = 1; row <= 256; row++) {
                        string color = row.ToString("X6");
                        string format = "0.00\"" + row + "\"";
                        for (int column = 1; column <= 2; column++) {
                            sheet.CellAt(row, column).SetValue(row);
                            sheet.CellBackground(row, column, color);
                            sheet.CellFontName(row, column, "Bounded Font");
                            sheet.FormatCell(row, column, format);
                        }
                    }
                    second.CellAt(1, 1).SetValue(7);
                    second.CellBackground(1, 1, "ABCDEF");
                    second.FormatCell(1, 1, "0.00\"other\"");
                }
                document.Save();
            }

            using (var package = SpreadsheetDocument.Open(path, false)) {
                Stylesheet styles = package.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
                Assert.Equal(259, styles.Fills!.ChildElements.Count);
                Assert.Equal(257, styles.NumberingFormats!.ChildElements.Count);
                Assert.True(styles.CellFormats!.ChildElements.Count < 1_300);
            }
            using (var reopened = ExcelDocument.Load(path)) {
                ExcelSheet sheet = reopened.Sheets[0];
                Assert.Equal("Bounded Font", sheet.CellAt(256, 2).GetStyle().FontName);
                Assert.Equal("0.00\"256\"", sheet.CellAt(256, 2).GetStyle().NumberFormatCode);
                Assert.Equal("FF000100", sheet.CellAt(256, 2).GetStyle().FillColorArgb);
                Assert.Equal("0.00\"other\"", reopened.Sheets[1].CellAt(1, 1).GetStyle().NumberFormatCode);
            }
        }
    }
}
