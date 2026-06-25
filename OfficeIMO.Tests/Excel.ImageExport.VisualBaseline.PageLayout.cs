using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class ExcelImageExportVisualBaselineTests {
        private const string PageLayoutBaselineName = "officeimo-excel-image-page-layout";

        [Fact]
        public void PageLayoutImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreatePageLayoutBaselineWorkbook();
            ExcelWorksheetImageExportOptions options = CreatePageLayoutBaselineOptions();

            IReadOnlyList<OfficeImageExportResult> pngResults = fixture.Sheet.ExportImages(OfficeImageExportFormat.Png, options);
            IReadOnlyList<OfficeImageExportResult> svgResults = fixture.Sheet.ExportImages(OfficeImageExportFormat.Svg, options);

            Assert.Equal(2, pngResults.Count);
            Assert.Equal(2, svgResults.Count);
            OfficeImageExportResult png = pngResults[1];
            OfficeImageExportResult svg = svgResults[1];
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.Equal("PageLayout!A3:D8", png.Source);
            Assert.Equal("PageLayout!A3:D8", svg.Source);
            Assert.Equal(1056, png.Width);
            Assert.Equal(872, png.Height);
            Assert.Equal(1056, svg.Width);
            Assert.Equal(872, svg.Height);
            Assert.Contains(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ManualPageBreaksSplit);
            Assert.Contains(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ManualPageBreaksSplit);
            Assert.Contains(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterFormattingApproximation);
            Assert.Contains(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterFormattingApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PrintTitlesUnsupported);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PrintTitlesUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupUnsupported);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupPaperSizeDefaulted);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupPaperSizeDefaulted);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("width=\"1056\"", svgText, StringComparison.Ordinal);
            Assert.Contains("height=\"872\"", svgText, StringComparison.Ordinal);
            Assert.Contains("Page 2 / 2 - PageLayout", svgText, StringComparison.Ordinal);
            Assert.Contains("Reviewed", svgText, StringComparison.Ordinal);
            Assert.Contains("Region", svgText, StringComparison.Ordinal);
            Assert.Contains("East", svgText, StringComparison.Ordinal);
            Assert.Contains("South", svgText, StringComparison.Ordinal);
            Assert.Contains("Global", svgText, StringComparison.Ordinal);
            Assert.Contains("Partner", svgText, StringComparison.Ordinal);
            AssertRasterBaseline(PageLayoutBaselineName + ".png", png.Bytes);
            AssertTextBaseline(PageLayoutBaselineName + ".svg", svgText);
        }

        [Fact]
        public void ApprovedPageLayoutBaselinesAreRenderableAndNonBlank() {
            string pngPath = Path.Combine(BaselineDirectory, PageLayoutBaselineName + ".png");
            string svgPath = Path.Combine(BaselineDirectory, PageLayoutBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreatePageLayoutBaselineWorkbook();
                ExcelWorksheetImageExportOptions options = CreatePageLayoutBaselineOptions();
                AssertRasterBaseline(PageLayoutBaselineName + ".png", fixture.Sheet.ExportImages(OfficeImageExportFormat.Png, options)[1].Bytes);
                AssertTextBaseline(PageLayoutBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(fixture.Sheet.ExportImages(OfficeImageExportFormat.Svg, options)[1].Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved page-layout PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved page-layout SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved page-layout PNG baseline is not a supported PNG file.");
            Assert.Equal(1056, image.Width);
            Assert.Equal(872, image.Height);
            int nonBackgroundPixels = VisualBaselineTestSupport.CountNonBackgroundPixels(image, OfficeColor.White);
            int headerPixels = CountPixelsNear(image, OfficeColor.FromRgb(15, 23, 42));
            int statusPixels = CountPixelsNear(image, OfficeColor.FromRgb(252, 228, 228));
            Assert.True(nonBackgroundPixels >= 10000, "Page-layout PNG baseline appears blank or nearly blank. Visible pixels: " + nonBackgroundPixels + ".");
            Assert.True(headerPixels >= 1000, "Page-layout PNG baseline is missing the dark worksheet/header chrome.");
            Assert.True(statusPixels >= 400, "Page-layout PNG baseline is missing the second-page status fill.");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Page 2 / 2 - PageLayout", svg, StringComparison.Ordinal);
            Assert.Contains("Reviewed", svg, StringComparison.Ordinal);
            Assert.Contains("Region", svg, StringComparison.Ordinal);
            Assert.Contains("East", svg, StringComparison.Ordinal);
            Assert.Contains("South", svg, StringComparison.Ordinal);
            Assert.Contains("Global", svg, StringComparison.Ordinal);
            Assert.Contains("Partner", svg, StringComparison.Ordinal);
        }

        private static ExcelBaselineFixture CreatePageLayoutBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelPageLayoutBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("PageLayout");

            string[] headers = { "Region", "Owner", "Status", "Score" };
            for (int column = 1; column <= headers.Length; column++) {
                sheet.CellValue(1, column, headers[column - 1]);
                sheet.CellAt(1, column).SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold().SetFontSize(12D);
                sheet.CellAlign(1, column, HorizontalAlignmentValues.Center);
                sheet.CellVerticalAlign(1, column, VerticalAlignmentValues.Center);
            }

            AddPageLayoutRow(sheet, 2, "North", "Avery", "Ready", 0.94D, "E7F6E7", "226B22");
            AddPageLayoutRow(sheet, 3, "West", "Morgan", "Ready", 0.89D, "E7F6E7", "226B22");
            AddPageLayoutRow(sheet, 4, "Central", "Quinn", "Watch", 0.77D, "FFF4CC", "7A4D00");
            AddPageLayoutRow(sheet, 5, "East", "Blake", "Watch", 0.71D, "FFF4CC", "7A4D00");
            AddPageLayoutRow(sheet, 6, "South", "Casey", "Risk", 0.82D, "FCE4E4", "9C0006");
            AddPageLayoutRow(sheet, 7, "Global", "Devon", "Ready", 0.91D, "E7F6E7", "226B22");
            AddPageLayoutRow(sheet, 8, "Partner", "Emery", "Risk", 0.64D, "FCE4E4", "9C0006");

            sheet.SetColumnWidth(1, 42D);
            sheet.SetColumnWidth(2, 38D);
            sheet.SetColumnWidth(3, 30D);
            sheet.SetColumnWidth(4, 26D);
            sheet.SetRowHeight(1, 32D);
            for (int row = 2; row <= 8; row++) {
                sheet.SetRowHeight(row, 50D);
            }

            for (int row = 1; row <= 8; row++) {
                for (int column = 1; column <= 4; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "94A3B8").SetFontSize(12D);
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            document.SetPrintTitles(sheet, firstRow: 1, lastRow: 1, firstCol: null, lastCol: null, save: false);
            sheet.SetOrientation(ExcelPageOrientation.Landscape);
            sheet.SetMargins(0.35D, 0.35D, 0.35D, 0.35D);
            sheet.SetPageSetup(fitToWidth: 1, fitToHeight: 0, paperSize: ExcelPaperSize.Letter);
            sheet.SetHeaderFooter(headerCenter: "&BPage &P / &N - &A", footerRight: "&IReviewed");
            sheet.AddManualRowPageBreak(2, save: false);
            return new ExcelBaselineFixture(document, sheet);
        }

        private static void AddPageLayoutRow(ExcelSheet sheet, int row, string region, string owner, string status, double score, string statusFill, string statusText) {
            sheet.CellValue(row, 1, region);
            sheet.CellValue(row, 2, owner);
            sheet.CellValue(row, 3, status);
            sheet.CellValue(row, 4, score);
            sheet.Range("A" + row.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":D" + row.ToString(System.Globalization.CultureInfo.InvariantCulture))
                .SetFillColor(row == 2 ? "F8FAFC" : "FFFFFF");
            sheet.CellAt(row, 3).SetFillColor(statusFill).SetFontColor(statusText).SetBold().SetFontSize(12D);
            sheet.CellAt(row, 4).Percent(0).SetBold().SetFontSize(12D);
            sheet.CellAlign(row, 4, HorizontalAlignmentValues.Right);
        }

        private static ExcelWorksheetImageExportOptions CreatePageLayoutBaselineOptions() =>
            new ExcelWorksheetImageExportOptions {
                Range = "A1:D8",
                SplitByManualPageBreaks = true,
                ShowGridlines = false,
                Scale = 1D,
                BackgroundColor = OfficeColor.White
            };
    }
}
