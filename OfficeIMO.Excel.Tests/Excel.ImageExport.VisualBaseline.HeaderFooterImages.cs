using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class ExcelImageExportVisualBaselineTests {
        private const string HeaderFooterImagesBaselineName = "officeimo-excel-image-header-footer-images";

        [Fact]
        public void HeaderFooterImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateHeaderFooterImagesBaselineWorkbook();
            ExcelWorksheetImageExportOptions options = CreateHeaderFooterImagesBaselineOptions();

            OfficeImageExportResult png = fixture.Sheet.ExportImages(OfficeImageExportFormat.Png, options)[1];
            OfficeImageExportResult svg = fixture.Sheet.ExportImages(OfficeImageExportFormat.Svg, options)[1];
            string svgText = Encoding.UTF8.GetString(svg.Bytes);

            Assert.Equal("HeaderFooterImages!A3:D4", png.Source);
            Assert.Equal("HeaderFooterImages!A3:D4", svg.Source);
            Assert.Equal(1056, png.Width);
            Assert.Equal(816, png.Height);
            Assert.Equal(1056, svg.Width);
            Assert.Equal(816, svg.Height);
            Assert.Contains(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterImageApproximation);
            Assert.Contains(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterImageApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("xl-header-footer-header-center-image", svgText, StringComparison.Ordinal);
            Assert.Contains("xl-header-footer-footer-right-image", svgText, StringComparison.Ordinal);
            Assert.Equal(2, svgText.Split(new[] { "data:image/png;base64," }, StringSplitOptions.None).Length - 1);
            Assert.DoesNotContain("&G", svgText, StringComparison.Ordinal);
            AssertRasterBaseline(HeaderFooterImagesBaselineName + ".png", png.Bytes);
            AssertTextBaseline(HeaderFooterImagesBaselineName + ".svg", svgText);
        }

        [Fact]
        public void ApprovedHeaderFooterImageBaselinesAreRenderableAndNonBlank() {
            string pngPath = Path.Combine(BaselineDirectory, HeaderFooterImagesBaselineName + ".png");
            string svgPath = Path.Combine(BaselineDirectory, HeaderFooterImagesBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateHeaderFooterImagesBaselineWorkbook();
                ExcelWorksheetImageExportOptions options = CreateHeaderFooterImagesBaselineOptions();
                AssertRasterBaseline(HeaderFooterImagesBaselineName + ".png", fixture.Sheet.ExportImages(OfficeImageExportFormat.Png, options)[1].Bytes);
                AssertTextBaseline(HeaderFooterImagesBaselineName + ".svg", Encoding.UTF8.GetString(fixture.Sheet.ExportImages(OfficeImageExportFormat.Svg, options)[1].Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved header/footer image PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved header/footer image SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved header/footer image PNG baseline is not a supported PNG file.");
            Assert.Equal(1056, image.Width);
            Assert.Equal(816, image.Height);
            int nonBackgroundPixels = VisualBaselineTestSupport.CountNonBackgroundPixels(image, OfficeColor.White);
            int headerLogoPixels = CountPixelsNear(image, OfficeColor.FromRgb(220, 38, 38));
            int footerLogoPixels = CountPixelsNear(image, OfficeColor.FromRgb(254, 240, 138));
            Assert.True(nonBackgroundPixels >= 10000, "Header/footer image PNG baseline appears blank or nearly blank. Visible pixels: " + nonBackgroundPixels + ".");
            Assert.True(headerLogoPixels >= 300, "Header/footer image PNG baseline is missing the red header image.");
            Assert.True(footerLogoPixels >= 120, "Header/footer image PNG baseline is missing the footer image.");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("xl-header-footer-header-center-image", svg, StringComparison.Ordinal);
            Assert.Contains("xl-header-footer-footer-right-image", svg, StringComparison.Ordinal);
            Assert.Equal(2, svg.Split(new[] { "data:image/png;base64," }, StringSplitOptions.None).Length - 1);
            Assert.Contains("Header/Footer Image Baseline", svg, StringComparison.Ordinal);
            Assert.Contains("Rendered on page 2", svg, StringComparison.Ordinal);
        }

        private static ExcelBaselineFixture CreateHeaderFooterImagesBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelHeaderFooterImagesBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("HeaderFooterImages");

            string[] headers = { "Area", "Owner", "State", "Score" };
            for (int column = 1; column <= headers.Length; column++) {
                sheet.CellValue(1, column, headers[column - 1]);
                sheet.CellAt(1, column).SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold().SetFontSize(12D);
                sheet.CellAlign(1, column, HorizontalAlignmentValues.Center);
                sheet.CellVerticalAlign(1, column, VerticalAlignmentValues.Center);
            }

            AddHeaderFooterImageRow(sheet, 2, "North", "Avery", "Ready", 0.94D, "E7F6E7", "166534");
            AddHeaderFooterImageRow(sheet, 3, "South", "Blake", "Review", 0.81D, "FEF3C7", "92400E");
            AddHeaderFooterImageRow(sheet, 4, "West", "Casey", "Risk", 0.68D, "FEE2E2", "991B1B");

            sheet.SetColumnWidth(1, 42D);
            sheet.SetColumnWidth(2, 38D);
            sheet.SetColumnWidth(3, 30D);
            sheet.SetColumnWidth(4, 26D);
            sheet.SetRowHeight(1, 32D);
            for (int row = 2; row <= 4; row++) {
                sheet.SetRowHeight(row, 58D);
            }

            for (int row = 1; row <= 4; row++) {
                for (int column = 1; column <= 4; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "94A3B8").SetFontSize(12D);
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            document.SetPrintTitles(sheet, firstRow: 1, lastRow: 1, firstCol: null, lastCol: null, save: false);
            sheet.SetOrientation(ExcelPageOrientation.Landscape);
            sheet.SetMargins(0.35D, 0.35D, 0.35D, 0.35D);
            sheet.SetPageSetup(fitToWidth: 1, fitToHeight: 0, paperSize: ExcelPaperSize.Letter);
            sheet.SetHeaderFooter(headerLeft: "&BHeader/Footer Image Baseline", footerCenter: "Rendered on page &P");
            sheet.SetHeaderImage(HeaderFooterPosition.Center, CreateHeaderFooterHeaderLogoPng(), "image/png", widthPoints: 72D, heightPoints: 24D);
            sheet.SetFooterImage(HeaderFooterPosition.Right, CreateHeaderFooterFooterLogoPng(), "image/png", widthPoints: 56D, heightPoints: 18D);
            sheet.AddManualRowPageBreak(2, save: false);
            return new ExcelBaselineFixture(document, sheet);
        }

        private static void AddHeaderFooterImageRow(ExcelSheet sheet, int row, string area, string owner, string state, double score, string stateFill, string stateText) {
            sheet.CellValue(row, 1, area);
            sheet.CellValue(row, 2, owner);
            sheet.CellValue(row, 3, state);
            sheet.CellValue(row, 4, score);
            sheet.Range("A" + row.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":D" + row.ToString(System.Globalization.CultureInfo.InvariantCulture))
                .SetFillColor(row == 2 ? "F8FAFC" : "FFFFFF");
            sheet.CellAt(row, 3).SetFillColor(stateFill).SetFontColor(stateText).SetBold().SetFontSize(12D);
            sheet.CellAt(row, 4).Percent(0).SetBold().SetFontSize(12D);
            sheet.CellAlign(row, 4, HorizontalAlignmentValues.Right);
        }

        private static ExcelWorksheetImageExportOptions CreateHeaderFooterImagesBaselineOptions() =>
            new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false,
                Scale = 1D,
                BackgroundColor = OfficeColor.White
            };

        private static byte[] CreateHeaderFooterHeaderLogoPng() {
            OfficeRasterImage image = new OfficeRasterImage(144, 48, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            canvas.FillRectangle(0, 0, 144, 48, OfficeColor.FromRgb(220, 38, 38));
            canvas.FillRectangle(6, 6, 132, 36, OfficeColor.White);
            canvas.FillRectangle(14, 14, 28, 20, OfficeColor.FromRgb(15, 23, 42));
            canvas.FillRectangle(50, 15, 72, 6, OfficeColor.FromRgb(220, 38, 38));
            canvas.FillRectangle(50, 27, 56, 5, OfficeColor.FromRgb(15, 23, 42));
            canvas.FillRectangle(112, 27, 10, 5, OfficeColor.FromRgb(220, 38, 38));
            return OfficePngWriter.Encode(image, OfficePngCompression.Stored);
        }

        private static byte[] CreateHeaderFooterFooterLogoPng() {
            OfficeRasterImage image = new OfficeRasterImage(112, 36, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            canvas.FillRectangle(0, 0, 112, 36, OfficeColor.FromRgb(37, 99, 235));
            canvas.FillRectangle(6, 6, 100, 24, OfficeColor.FromRgb(248, 250, 252));
            canvas.FillRectangle(14, 13, 64, 10, OfficeColor.FromRgb(254, 240, 138));
            canvas.FillRectangle(84, 13, 10, 10, OfficeColor.FromRgb(37, 99, 235));
            canvas.DrawRectangle(3, 3, 106, 30, OfficeColor.FromRgb(30, 64, 175), 2);
            return OfficePngWriter.Encode(image, OfficePngCompression.Stored);
        }
    }
}
