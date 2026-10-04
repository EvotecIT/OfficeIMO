using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
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
