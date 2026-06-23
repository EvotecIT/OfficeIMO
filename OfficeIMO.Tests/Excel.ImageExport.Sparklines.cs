using DocumentFormat.OpenXml.Office2010.Excel;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelImageExportSparklineTests {
        [Fact]
        public void ExcelRange_ImageExportRendersLineSparklinesInVisibleRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Sparklines");
            sheet.CellValue(1, 1, "Metric");
            sheet.CellValue(2, 1, "Revenue");
            sheet.CellValue(2, 2, 10);
            sheet.CellValue(2, 3, 18);
            sheet.CellValue(2, 4, 14);
            sheet.CellValue(2, 6, 7);
            sheet.CellValue(2, 7, 9);
            sheet.AddSparklines("B2:D2", "E2:E2", displayMarkers: true, seriesColor: "#2563EB");
            sheet.AddSparklines("F2:G2", "H2:H2", displayMarkers: true, seriesColor: "#DC2626");

            ExcelRange range = sheet.Range("A1:E2");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            ExcelVisualSparkline visual = Assert.Single(snapshot.Sparklines);
            Assert.Equal("Sparklines!E2", visual.Source);
            Assert.Equal(new[] { 10D, 18D, 14D }, visual.Values);
            Assert.Contains(snapshot.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.SparklineRenderingApproximation && diagnostic.Source == "Sparklines!E2");
            Assert.Contains(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.SparklineRenderingApproximation && diagnostic.Source == "Sparklines!E2");
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code == "ExcelSparklineUnsupported");
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == "ExcelSparklineUnsupported");
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Source == "Sparklines!H2");
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Source == "Sparklines!H2");
            Assert.Contains("<polyline", svg, StringComparison.Ordinal);
            Assert.Contains("<circle", svg, StringComparison.Ordinal);
            Assert.Contains("#2563EB", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(CountBluePixels(rendered!) > 0);
        }

        [Fact]
        public void ExcelRange_ImageExportRendersColumnAndWinLossSparklines() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Types");
            sheet.CellValue(1, 1, 8);
            sheet.CellValue(1, 2, -4);
            sheet.CellValue(1, 3, 12);
            sheet.CellValue(2, 1, 1);
            sheet.CellValue(2, 2, -1);
            sheet.CellValue(2, 3, 1);
            sheet.AddSparklines("A1:C1", "D1", SparklineTypeValues.Column, displayNegative: true, displayAxis: true, seriesColor: "#16A34A", negativeColor: "#DC2626", axisColor: "#6B7280");
            sheet.AddSparklines("A2:C2", "D2", SparklineTypeValues.Stacked, displayNegative: true, displayAxis: true, seriesColor: "#0EA5E9", negativeColor: "#DC2626", axisColor: "#6B7280");

            ExcelRange range = sheet.Range("A1:D2");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Equal(2, snapshot.Sparklines.Count);
            Assert.Contains(snapshot.Sparklines, sparkline => sparkline.Kind == "column");
            Assert.Contains(snapshot.Sparklines, sparkline => sparkline.Kind == "stacked");
            Assert.Equal(2, snapshot.Diagnostics.Count(diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.SparklineRenderingApproximation));
            Assert.Contains("<rect", svg, StringComparison.Ordinal);
            Assert.Contains("#16A34A", svg, StringComparison.Ordinal);
            Assert.Contains("#DC2626", svg, StringComparison.Ordinal);
            Assert.Contains("#6B7280", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportReportsExternalSparklineRanges() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet data = document.AddWorkSheet("Data");
            data.CellValue(1, 1, 2);
            data.CellValue(1, 2, 4);
            data.CellValue(1, 3, 8);
            ExcelSheet summary = document.AddWorkSheet("Summary");
            summary.AddSparklines("'Data'!A1:C1", "A1", displayMarkers: true);

            ExcelRange range = summary.Range("A1:A1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png);

            Assert.Empty(snapshot.Sparklines);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(snapshot.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.SparklineExternalRangeUnsupported);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("Summary!A1", diagnostic.Source);
            Assert.Contains(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.SparklineExternalRangeUnsupported && item.Source == "Summary!A1");
        }

        private static int CountBluePixels(OfficeRasterImage image) {
            int count = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor color = image.GetPixel(x, y);
                    if (color.B > 150 && color.R < 120 && color.G < 150) {
                        count++;
                    }
                }
            }

            return count;
        }
    }
}
