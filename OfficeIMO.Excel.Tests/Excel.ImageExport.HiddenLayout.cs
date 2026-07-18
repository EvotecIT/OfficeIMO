using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelImageExportHiddenLayoutTests {
        [Fact]
        public void ExcelRange_ImageExportOmitsHiddenRowsColumnsAndReportsHiddenAnchoredObjects() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Hidden");
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

        [Fact]
        public void ExcelRange_ImageExportOmitsDefaultHiddenRowsUnlessExplicitlyVisible() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("DefaultHidden");
            sheet.CellValue(1, 1, "Hidden by default");
            sheet.CellValue(2, 1, "Visible override");
            sheet.CellValue(3, 1, "Hidden by default too");
            sheet.SetDefaultRowHeight(18D, hidden: true);
            sheet.SetRowHeight(2, 18D);

            ExcelRange range = sheet.Range("A1:A3");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });
            ExcelRangeVisualSnapshot includeHidden = range.CreateVisualSnapshot(new ExcelImageExportOptions { IncludeHidden = true, ShowGridlines = false });

            ExcelVisualRow row = Assert.Single(snapshot.Rows);
            Assert.Equal(2, row.Index);
            ExcelVisualCell cell = Assert.Single(snapshot.Cells);
            Assert.Equal("Visible override", cell.Text);
            AssertDiagnostic(snapshot.Diagnostics, ExcelImageExportDiagnosticCodes.HiddenRowsOmitted, "DefaultHidden!A1:A3");
            Assert.Equal(new[] { 1, 2, 3 }, includeHidden.Rows.Select(item => item.Index).ToArray());
        }

        [Fact]
        public void ExcelWorksheet_DefaultImageExportSkipsHiddenImageAnchorsWhenExpandingUsedRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("HiddenAnchor");
            sheet.CellValue(1, 1, "Visible");
            sheet.SetColumnHidden(10, true);
            sheet.AddImage(1, 10, CreateSolidPng(24, 18, OfficeColor.FromRgb(220, 38, 38)), "image/png", widthPixels: 24, heightPixels: 18, name: "HiddenLogo");

            OfficeImageExportResult defaultResult = sheet.ExportImage(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions { ShowGridlines = false });
            OfficeImageExportResult includeHiddenResult = sheet.ExportImage(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions { IncludeHidden = true, ShowGridlines = false });

            Assert.Equal("HiddenAnchor!A1:A1", defaultResult.Source);
            Assert.Equal("HiddenAnchor!A1:J1", includeHiddenResult.Source);
            Assert.True(includeHiddenResult.Width > defaultResult.Width, $"Expected including the hidden anchored image to expand the default range. default={defaultResult.Width}, includeHidden={includeHiddenResult.Width}");
        }

        [Fact]
        public void ExcelWorksheet_DefaultImageExportSkipsDefaultHiddenImageAnchorsWhenExpandingUsedRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("DefaultHiddenAnchor");
            sheet.CellValue(1, 1, "Visible");
            sheet.SetDefaultRowHeight(18D, hidden: true);
            sheet.SetRowHeight(1, 18D);
            sheet.AddImage(10, 1, CreateSolidPng(24, 18, OfficeColor.FromRgb(220, 38, 38)), "image/png", widthPixels: 24, heightPixels: 18, name: "DefaultHiddenLogo");

            OfficeImageExportResult defaultResult = sheet.ExportImage(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions { ShowGridlines = false });
            OfficeImageExportResult includeHiddenResult = sheet.ExportImage(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions { IncludeHidden = true, ShowGridlines = false });

            Assert.Equal("DefaultHiddenAnchor!A1:A1", defaultResult.Source);
            Assert.Equal("DefaultHiddenAnchor!A1:A10", includeHiddenResult.Source);
            Assert.True(includeHiddenResult.Height > defaultResult.Height, $"Expected including the default-hidden anchored image to expand the default range. default={defaultResult.Height}, includeHidden={includeHiddenResult.Height}");
        }

        [Fact]
        public void ExcelWorkbook_DefaultImageExportSkipsHiddenSheetsUnlessRequested() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet visible = document.AddWorksheet("Visible");
            ExcelSheet hidden = document.AddWorksheet("Hidden");
            ExcelSheet veryHidden = document.AddWorksheet("VeryHidden");
            visible.CellValue(1, 1, "Visible");
            hidden.CellValue(1, 1, "Hidden");
            veryHidden.CellValue(1, 1, "Very hidden");
            hidden.SetHidden(true);
            veryHidden.SetVeryHidden(true);

            IReadOnlyList<OfficeImageExportResult> defaultResults = document.ExportImages(OfficeImageExportFormat.Svg);
            IReadOnlyList<OfficeImageExportResult> explicitResults = document.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorkbookImageExportOptions {
                SheetNames = new[] { "Hidden" }
            });
            IReadOnlyList<OfficeImageExportResult> includeHiddenResults = document.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorkbookImageExportOptions {
                IncludeHiddenSheets = true
            });

            Assert.Equal(new[] { "Visible" }, defaultResults.Select(result => result.Name).ToArray());
            OfficeImageExportResult explicitResult = Assert.Single(explicitResults);
            Assert.Equal("Hidden", explicitResult.Name);
            Assert.Equal(new[] { "Visible", "Hidden", "VeryHidden" }, includeHiddenResults.Select(result => result.Name).ToArray());
        }

        [Fact]
        public void ExcelWorksheet_DefaultImageExportUsesVisibleExtentsAcrossHiddenColumns() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("VisibleExtent");
            sheet.CellValue(1, 1, "Visible");
            sheet.SetColumnHidden(2, true);
            sheet.AddImage(1, 1, CreateSolidPng(120, 18, OfficeColor.FromRgb(37, 99, 235)), "image/png", widthPixels: 120, heightPixels: 18, name: "VisibleLogo");

            OfficeImageExportResult result = sheet.ExportImage(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions { ShowGridlines = false });

            Assert.Equal("VisibleExtent!A1:C1", result.Source);
            Assert.True(result.Width > 100, $"Expected the visible canvas to extend beyond column A when hidden column B has no visible width. width={result.Width}");
        }

        [Fact]
        public void ExcelWorksheet_DefaultImageExportIncludesImageAnchorOffsetsWhenExpandingUsedRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("OffsetImage");
            sheet.CellValue(1, 1, "Visible");
            sheet.SetColumnWidth(1, 8);
            sheet.SetColumnWidth(2, 8);
            sheet.AddImage(1, 1, CreateSolidPng(70, 18, OfficeColor.FromRgb(37, 99, 235)), "image/png", widthPixels: 70, heightPixels: 18, offsetXPixels: 58, name: "OffsetLogo");

            OfficeImageExportResult result = sheet.ExportImage(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions { ShowGridlines = false });

            Assert.Equal("OffsetImage!A1:C1", result.Source);
        }

        [Fact]
        public void ExcelWorksheet_DefaultPngExportExpandsRangeForVisibleImageFallbacks() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("UnsupportedImage");
            sheet.CellValue(1, 1, "Visible");
            sheet.AddImage(12, 8, CreateMinimalJpegHeader(), "image/jpeg", widthPixels: 64, heightPixels: 32, name: "JpegOutside");

            OfficeImageExportResult png = sheet.ExportImage(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions { ShowGridlines = false });
            ExcelRangeVisualSnapshot snapshot = sheet.CreateVisualSnapshot(new ExcelWorksheetImageExportOptions { ShowGridlines = false });

            Assert.Equal("UnsupportedImage!A1:H13", png.Source);
            Assert.Equal("A1:H13", snapshot.Range);
            Assert.Contains(
                png.Diagnostics,
                diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.Contains(snapshot.Images, image => image.Name == "JpegOutside" && image.DetectedFormat == OfficeImageFormat.Jpeg);
        }

        [Fact]
        public void ExcelRange_ImageExportEvaluatesConditionalFormattingForMergeOriginOutsideSelectedRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("MergeCf");
            sheet.CellValue(1, 1, 20);
            sheet.MergeRange("A1:B1");
            sheet.AddConditionalRule("A1:B1", DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingOperatorValues.GreaterThan, "10", fillColor: "C6EFCE");

            ExcelRangeVisualSnapshot snapshot = sheet.Range("B1:C1").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });

            ExcelVisualCell origin = Assert.Single(snapshot.Cells, cell => cell.Row == 1 && cell.Column == 1);
            Assert.Equal("FFC6EFCE", origin.Style.FillColorArgb);
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

        private static byte[] CreateMinimalJpegHeader() {
            return new byte[] {
                0xFF, 0xD8,
                0xFF, 0xE0, 0x00, 0x10,
                0x4A, 0x46, 0x49, 0x46, 0x00,
                0x01, 0x01, 0x00,
                0x00, 0x01, 0x00, 0x01,
                0x00, 0x00,
                0xFF, 0xC0, 0x00, 0x11,
                0x08,
                0x00, 0x01,
                0x00, 0x01,
                0x03,
                0x01, 0x11, 0x00,
                0x02, 0x11, 0x00,
                0x03, 0x11, 0x00,
                0xFF, 0xD9
            };
        }
    }
}
