using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class ExcelImageExportTests {
        [Fact]
        public void ExcelRange_ImageExportAppliesTimePeriodConditionalFillsWithReferenceDate() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Dates");
            sheet.CellValue(1, 1, new DateTime(2026, 6, 24));
            sheet.CellValue(2, 1, new DateTime(2026, 6, 23));
            sheet.CellValue(3, 1, new DateTime(2026, 6, 25));
            sheet.CellValue(1, 2, new DateTime(2026, 6, 24));
            sheet.CellValue(2, 2, new DateTime(2026, 6, 18));
            sheet.CellValue(3, 2, new DateTime(2026, 6, 17));
            sheet.Range("A1:B3").DateTime("yyyy-mm-dd");
            sheet.SetColumnWidth(1, 12);
            sheet.SetColumnWidth(2, 12);
            sheet.SetRowHeight(1, 30);
            sheet.SetRowHeight(2, 30);
            sheet.SetRowHeight(3, 30);
            sheet.AddConditionalTimePeriodRule("A1:A3", TimePeriodValues.Today, fillColor: "C6EFCE");
            sheet.AddConditionalTimePeriodRule("B1:B3", TimePeriodValues.Last7Days, fillColor: "DBEAFE");

            var options = new ExcelImageExportOptions {
                ShowGridlines = false,
                ConditionalFormattingDate = new DateTime(2026, 6, 24)
            };
            ExcelRange range = sheet.Range("A1:B3");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.Equal("FFC6EFCE", snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 1).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 1).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 1).Style.FillColorArgb);
            Assert.Equal("FFDBEAFE", snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 2).Style.FillColorArgb);
            Assert.Equal("FFDBEAFE", snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 2).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 2).Style.FillColorArgb);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ConditionalTimePeriodUnsupported);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ConditionalRuleUnsupported);
            Assert.Contains("#C6EFCE", svg, StringComparison.Ordinal);
            Assert.Contains("#DBEAFE", svg, StringComparison.Ordinal);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            ExcelVisualCell todayCell = snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 1);
            ExcelVisualCell last7DaysCell = snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 2);
            AssertCellHasColorPatch(rendered!, todayCell, OfficeColor.FromRgb(198, 239, 206), tolerance: 3);
            AssertCellHasColorPatch(rendered!, last7DaysCell, OfficeColor.FromRgb(219, 234, 254), tolerance: 3);

            ExcelConditionalFormattingInfo today = Assert.Single(sheet.GetConditionalFormattingRules("A1:A3"), rule => rule.Type == "TimePeriod");
            ExcelConditionalFormattingInfo last7Days = Assert.Single(sheet.GetConditionalFormattingRules("B1:B3"), rule => rule.Type == "TimePeriod");
            Assert.Equal(nameof(TimePeriodValues.Today), today.TimePeriod);
            Assert.Equal(nameof(TimePeriodValues.Last7Days), last7Days.TimePeriod);
            Assert.Equal("FFC6EFCE", today.DifferentialFillColorArgb);
            Assert.Equal("FFDBEAFE", last7Days.DifferentialFillColorArgb);
        }

        private static void AssertCellHasColorPatch(OfficeRasterImage image, ExcelVisualCell cell, OfficeColor expected, int tolerance) {
            int matches = 0;
            int left = Math.Max(0, (int)Math.Ceiling(cell.X) + 2);
            int top = Math.Max(0, (int)Math.Ceiling(cell.Y) + 2);
            int right = Math.Min(image.Width - 1, (int)Math.Floor(cell.X + cell.Width) - 2);
            int bottom = Math.Min(image.Height - 1, (int)Math.Floor(cell.Y + cell.Height) - 2);
            for (int y = top; y <= bottom; y++) {
                for (int x = left; x <= right; x++) {
                    OfficeColor actual = image.GetPixel(x, y);
                    if (Math.Abs(actual.R - expected.R) <= tolerance &&
                        Math.Abs(actual.G - expected.G) <= tolerance &&
                        Math.Abs(actual.B - expected.B) <= tolerance) {
                        matches++;
                    }
                }
            }

            Assert.True(matches >= 16, $"Expected at least 16 pixels near {expected} inside cell {cell.Row},{cell.Column}, found {matches}.");
        }
    }
}
