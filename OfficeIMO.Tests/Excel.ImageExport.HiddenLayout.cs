using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelImageExportHiddenLayoutTests {
        [Fact]
        public void ExcelRange_ImageExportOmitsHiddenRowsColumnsAndReportsHiddenAnchoredObjects() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Hidden");
            sheet.CellValue(1, 1, "Visible A");
            sheet.CellValue(1, 2, "Hidden B");
            sheet.CellValue(1, 3, "Visible C");
            sheet.CellValue(2, 1, "Hidden Row");
            sheet.CellValue(2, 2, 10);
            sheet.CellValue(2, 3, 20);
            sheet.CellValue(3, 1, "Visible Row");
            sheet.CellValue(3, 2, 30);
            sheet.CellValue(3, 3, 40);
            sheet.SetColumnWidth(1, 12);
            sheet.SetColumnWidth(2, 12);
            sheet.SetColumnWidth(3, 12);
            sheet.SetRowHeight(1, 24);
            sheet.SetRowHeight(2, 24);
            sheet.SetRowHeight(3, 24);
            sheet.SetColumnHidden(2, true);
            sheet.SetRowHidden(2, true);
            sheet.AddImage(1, 2, CreateSolidPng(12, 10, OfficeColor.FromRgb(220, 38, 38)), "image/png", widthPixels: 12, heightPixels: 10, name: "HiddenLogo");
            sheet.AddChartFromRange("A1:C3", row: 2, column: 3, widthPixels: 160, heightPixels: 90, type: ExcelChartType.ColumnClustered, title: "Hidden Chart");

            ExcelRange range = sheet.Range("A1:C3");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
            ExcelRangeVisualSnapshot includeHidden = range.CreateVisualSnapshot(new ExcelImageExportOptions { IncludeHidden = true });
            OfficeImageExportResult includeHiddenPng = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { IncludeHidden = true, ShowGridlines = false });

            Assert.Equal(new[] { 1, 3 }, snapshot.Columns.Select(column => column.Index).ToArray());
            Assert.Equal(new[] { 1, 3 }, snapshot.Rows.Select(row => row.Index).ToArray());
            Assert.DoesNotContain(snapshot.Cells, cell => cell.Text == "Hidden B" || cell.Text == "Hidden Row");
            Assert.Empty(snapshot.Images);
            Assert.Empty(snapshot.Charts);
            AssertDiagnostic(snapshot.Diagnostics, ExcelImageExportDiagnosticCodes.HiddenRowsOmitted, "Hidden!A1:C3");
            AssertDiagnostic(snapshot.Diagnostics, ExcelImageExportDiagnosticCodes.HiddenColumnsOmitted, "Hidden!A1:C3");
            AssertDiagnostic(snapshot.Diagnostics, ExcelImageExportDiagnosticCodes.ImageAnchorHidden, "Hidden!HiddenLogo");
            AssertDiagnosticSourcePrefix(snapshot.Diagnostics, ExcelImageExportDiagnosticCodes.ChartAnchorHidden, "Hidden!Chart");
            AssertDiagnostic(png.Diagnostics, ExcelImageExportDiagnosticCodes.HiddenRowsOmitted, "Hidden!A1:C3");
            AssertDiagnostic(png.Diagnostics, ExcelImageExportDiagnosticCodes.HiddenColumnsOmitted, "Hidden!A1:C3");
            AssertDiagnostic(png.Diagnostics, ExcelImageExportDiagnosticCodes.ImageAnchorHidden, "Hidden!HiddenLogo");
            AssertDiagnosticSourcePrefix(png.Diagnostics, ExcelImageExportDiagnosticCodes.ChartAnchorHidden, "Hidden!Chart");

            Assert.Equal(new[] { 1, 2, 3 }, includeHidden.Columns.Select(column => column.Index).ToArray());
            Assert.Equal(new[] { 1, 2, 3 }, includeHidden.Rows.Select(row => row.Index).ToArray());
            Assert.Contains(includeHidden.Cells, cell => cell.Text == "Hidden B");
            Assert.Contains(includeHidden.Cells, cell => cell.Text == "Hidden Row");
            Assert.Single(includeHidden.Images);
            Assert.Single(includeHidden.Charts);
            Assert.DoesNotContain(includeHidden.Diagnostics, diagnostic => diagnostic.Code.StartsWith("ExcelHidden", StringComparison.Ordinal));
            Assert.DoesNotContain(includeHidden.Diagnostics, diagnostic => diagnostic.Code.EndsWith("AnchorHidden", StringComparison.Ordinal));

            OfficeImageInfo defaultInfo = OfficeImageReader.Identify(png.Bytes);
            OfficeImageInfo includeHiddenInfo = OfficeImageReader.Identify(includeHiddenPng.Bytes);
            Assert.True(includeHiddenInfo.Width > defaultInfo.Width, $"Expected including hidden columns to make the PNG wider. default={defaultInfo.Width}, includeHidden={includeHiddenInfo.Width}");
            Assert.True(includeHiddenInfo.Height > defaultInfo.Height, $"Expected including hidden rows to make the PNG taller. default={defaultInfo.Height}, includeHidden={includeHiddenInfo.Height}");
        }

        private static void AssertDiagnostic(IReadOnlyList<OfficeImageExportDiagnostic> diagnostics, string code, string source) {
            OfficeImageExportDiagnostic diagnostic = Assert.Single(diagnostics, item => item.Code == code);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal(source, diagnostic.Source);
        }

        private static void AssertDiagnosticSourcePrefix(IReadOnlyList<OfficeImageExportDiagnostic> diagnostics, string code, string sourcePrefix) {
            OfficeImageExportDiagnostic diagnostic = Assert.Single(diagnostics, item => item.Code == code);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.StartsWith(sourcePrefix, diagnostic.Source, StringComparison.Ordinal);
        }

        private static byte[] CreateSolidPng(int width, int height, OfficeColor color) {
            OfficeRasterImage image = new OfficeRasterImage(width, height, OfficeColor.Transparent);
            image.Fill(color);
            return OfficePngWriter.Encode(image);
        }
    }
}
