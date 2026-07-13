using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class ExcelImageExportVisualBaselineTests {
        private const string TextSpillBaselineName = "officeimo-excel-image-text-spill";

        [Fact]
        public void TextSpillImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateTextSpillBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:E5");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && item.Source == "Spill!A2");
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && item.Source == "Spill!A3");
            Assert.Contains("Plain text spills through blanks", svgText, StringComparison.Ordinal);
            Assert.Contains("Rich", svgText, StringComparison.Ordinal);
            Assert.Contains("spill keeps runs", svgText, StringComparison.Ordinal);
            Assert.Contains("Stop", svgText, StringComparison.Ordinal);
            Assert.Contains("font-weight=\"700\"", svgText, StringComparison.Ordinal);
            Assert.Contains("font-style=\"italic\"", svgText, StringComparison.Ordinal);
            AssertRasterBaseline(TextSpillBaselineName + ".png", png.Bytes);
            AssertTextBaseline(TextSpillBaselineName + ".svg", svgText);
        }

        [Fact]
        public void ApprovedTextSpillBaselinesAreRenderableAndNonBlank() {
            string pngPath = Path.Combine(BaselineDirectory, TextSpillBaselineName + ".png");
            string svgPath = Path.Combine(BaselineDirectory, TextSpillBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateTextSpillBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:E5");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(TextSpillBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(TextSpillBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved text-spill PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved text-spill SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved text-spill PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 600, "Text-spill PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 180, "Text-spill PNG baseline height is unexpectedly small.");
            int nonBackgroundPixels = VisualBaselineTestSupport.CountNonBackgroundPixels(image, OfficeColor.White);
            Assert.True(nonBackgroundPixels >= 1400, "Text-spill PNG baseline appears blank or nearly blank. Visible pixels: " + nonBackgroundPixels + ".");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Excel Text Spill", svg, StringComparison.Ordinal);
            Assert.Contains("Plain text spills through blanks", svg, StringComparison.Ordinal);
            Assert.Contains("Rich", svg, StringComparison.Ordinal);
            Assert.Contains("Stop", svg, StringComparison.Ordinal);
            Assert.Contains("<clipPath", svg, StringComparison.Ordinal);
        }

        private static ExcelBaselineFixture CreateTextSpillBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Spill");

            sheet.CellValue(1, 1, "Excel Text Spill");
            sheet.Range("A1:E1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellValue(2, 1, "Plain text spills through blanks");
            sheet.CellValue(2, 4, "Stop");
            sheet.CellValue(2, 5, "occupied cell blocks");
            sheet.CellAt(3, 1).SetRichText(
                new ExcelRichTextRun("Rich") { Bold = true, FontColor = "0F766E", FontSize = 12D },
                new ExcelRichTextRun(" spill keeps runs") { Italic = true, FontColor = "7C3AED", FontSize = 12D });
            sheet.CellValue(3, 4, "Stop");
            sheet.CellValue(3, 5, "rich text path");
            sheet.CellValue(4, 1, "Wrapped\ncell text");
            sheet.WrapCells(4, 1, 1);
            sheet.CellValue(4, 4, "No spill");
            sheet.CellValue(5, 1, "Centered");
            sheet.CellAlign(5, 1, HorizontalAlignmentValues.Center);
            sheet.CellValue(5, 4, "Policy");

            sheet.SetColumnWidth(1, 9);
            sheet.SetColumnWidth(2, 16);
            sheet.SetColumnWidth(3, 16);
            sheet.SetColumnWidth(4, 10);
            sheet.SetColumnWidth(5, 18);
            sheet.SetRowHeight(1, 30);
            sheet.SetRowHeight(2, 28);
            sheet.SetRowHeight(3, 28);
            sheet.SetRowHeight(4, 42);
            sheet.SetRowHeight(5, 28);

            for (int row = 2; row <= 5; row++) {
                sheet.CellAt(row, 1).SetFillColor("F8FAFC");
                sheet.CellAt(row, 4).SetFillColor("E2E8F0").SetBold();
                sheet.CellAt(row, 5).SetFontColor("475569");
            }

            for (int row = 1; row <= 5; row++) {
                for (int column = 1; column <= 5; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                }
            }

            return new ExcelBaselineFixture(document, sheet);
        }
    }
}
