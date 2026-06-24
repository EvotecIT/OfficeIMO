using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class ExcelImageExportVisualBaselineTests {
        private const string ChartAxisLabelsBaselineName = "officeimo-excel-image-chart-axis-labels";

        [Fact]
        public void ChartAxisLabelsImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateChartAxisLabelsBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:H10");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartAxisNumberFormatApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains(">500<", svgText, StringComparison.Ordinal);
            Assert.Contains(">1,000<", svgText, StringComparison.Ordinal);
            Assert.Contains(">1,500<", svgText, StringComparison.Ordinal);
            AssertRasterBaseline(ChartAxisLabelsBaselineName + ".png", png.Bytes);
            AssertTextBaseline(ChartAxisLabelsBaselineName + ".svg", svgText);
        }

        [Fact]
        public void ApprovedChartAxisLabelsBaselinesAreRenderableAndNonBlank() {
            using ExcelBaselineFixture fixture = CreateChartAxisLabelsBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:H10");
            ExcelImageExportOptions options = CreateBaselineOptions();
            string pngPath = Path.Combine(BaselineDirectory, ChartAxisLabelsBaselineName + ".png");
            string svgPath = Path.Combine(BaselineDirectory, ChartAxisLabelsBaselineName + ".svg");

            Assert.True(File.Exists(pngPath), "Approved chart-axis PNG baseline is missing.");
            Assert.True(File.Exists(svgPath), "Approved chart-axis SVG baseline is missing.");
            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved chart-axis PNG baseline is not a supported PNG file.");
            Assert.Equal(range.ExportImage(OfficeImageExportFormat.Png, options).Width, image.Width);
            Assert.Equal(range.ExportImage(OfficeImageExportFormat.Png, options).Height, image.Height);
            Assert.True(VisualBaselineTestSupport.CountNonBackgroundPixels(image, OfficeColor.White) > 1200, "Approved chart-axis PNG baseline should contain a visible rendered chart.");
            string svg = File.ReadAllText(svgPath);
            Assert.Contains("Axis Label Density", svg, StringComparison.Ordinal);
            Assert.Contains(">1,500<", svg, StringComparison.Ordinal);
        }

        private static ExcelBaselineFixture CreateChartAxisLabelsBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ChartAxisLabels");

            sheet.CellValue(1, 1, "Chart Axis Label Density");
            sheet.Range("A1:H1").Merge();
            sheet.Range("A1:H1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);

            sheet.CellValue(2, 1, "Month");
            sheet.CellValue(2, 2, "Actual");
            sheet.CellValue(3, 1, "Jan");
            sheet.CellValue(3, 2, 1200);
            sheet.CellValue(4, 1, "Feb");
            sheet.CellValue(4, 2, 1800);
            sheet.CellValue(5, 1, "Mar");
            sheet.CellValue(5, 2, 1600);
            sheet.CellValue(6, 1, "Apr");
            sheet.CellValue(6, 2, 2000);
            sheet.Range("A2:B2").SetFillColor("E2E8F0").SetBold();
            sheet.Range("A3:B6").SetFillColor("F8FAFC");

            ExcelChart chart = sheet.AddChartFromRange("A2:B6", row: 2, column: 4, widthPixels: 340, heightPixels: 220, type: ExcelChartType.ColumnClustered, title: "Axis Label Density");
            chart.SetValueAxisNumberFormat("#,##0");
            chart.SetLegend(LegendPositionValues.Bottom);

            for (int column = 1; column <= 8; column++) {
                sheet.SetColumnWidth(column, column <= 2 ? 13 : 11);
            }

            sheet.SetRowHeight(1, 28);
            for (int row = 2; row <= 10; row++) {
                sheet.SetRowHeight(row, 24);
            }

            for (int row = 1; row <= 10; row++) {
                for (int column = 1; column <= 8; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            return new ExcelBaselineFixture(document, sheet);
        }
    }
}
