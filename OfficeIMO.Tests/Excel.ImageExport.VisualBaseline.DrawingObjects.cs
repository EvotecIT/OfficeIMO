using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class ExcelImageExportVisualBaselineTests {
        private const string RotatedPresetDrawingObjectBaselineName = "officeimo-excel-image-rotated-preset-drawing-object";
        private const string RotatedShapeTextBaselineName = "officeimo-excel-image-rotated-shape-text";
        private const string AlignedShapeTextBaselineName = "officeimo-excel-image-aligned-shape-text";
        private const string VerticalShapeTextBaselineName = "officeimo-excel-image-vertical-shape-text";

        [Fact]
        public void RotatedPresetDrawingObjectImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateRotatedPresetDrawingObjectBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:F8");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("matrix(", svgText, StringComparison.Ordinal);
            Assert.Contains("#FDBA74", svgText, StringComparison.Ordinal);
            Assert.Contains("#EA580C", svgText, StringComparison.Ordinal);
            AssertRasterBaseline(RotatedPresetDrawingObjectBaselineName + ".png", png.Bytes);
            AssertTextBaseline(RotatedPresetDrawingObjectBaselineName + ".svg", svgText);
        }

        [Fact]
        public void ApprovedRotatedPresetDrawingObjectBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, RotatedPresetDrawingObjectBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, RotatedPresetDrawingObjectBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateRotatedPresetDrawingObjectBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:F8");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(RotatedPresetDrawingObjectBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(RotatedPresetDrawingObjectBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved rotated preset drawing-object PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved rotated preset drawing-object SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved rotated preset drawing-object PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 600, "Rotated preset drawing-object PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 300, "Rotated preset drawing-object PNG baseline height is unexpectedly small.");
            int fillPixels = CountPixelsNear(image, OfficeColor.FromRgb(253, 186, 116));
            Assert.True(fillPixels > 1200, "Rotated preset drawing-object PNG baseline does not contain enough visible heart fill pixels.");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("<path", svg, StringComparison.Ordinal);
            Assert.Contains("matrix(", svg, StringComparison.Ordinal);
            Assert.Contains("#FDBA74", svg, StringComparison.Ordinal);
            Assert.Contains("#EA580C", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void RotatedShapeTextImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateRotatedShapeTextBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:F8");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation);
            Assert.Single(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("Rotated label", svgText, StringComparison.Ordinal);
            Assert.Contains("transform=\"rotate(24", svgText, StringComparison.Ordinal);
            Assert.Contains("#DBEAFE", svgText, StringComparison.Ordinal);
            Assert.Contains("#2563EB", svgText, StringComparison.Ordinal);
            AssertDiagnosticsBaseline(RotatedShapeTextBaselineName + ".diagnostics.txt", png.Diagnostics);
            AssertRasterBaseline(RotatedShapeTextBaselineName + ".png", png.Bytes);
            AssertTextBaseline(RotatedShapeTextBaselineName + ".svg", svgText);
        }

        [Fact]
        public void ApprovedRotatedShapeTextBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, RotatedShapeTextBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, RotatedShapeTextBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateRotatedShapeTextBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:F8");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(RotatedShapeTextBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(RotatedShapeTextBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved rotated shape-text PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved rotated shape-text SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved rotated shape-text PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 600, "Rotated shape-text PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 300, "Rotated shape-text PNG baseline height is unexpectedly small.");
            int fillPixels = CountPixelsNear(image, OfficeColor.FromRgb(219, 234, 254));
            int textPixels = CountPixelsNear(image, OfficeColor.FromRgb(31, 41, 55));
            Assert.True(fillPixels > 1200, "Rotated shape-text PNG baseline does not contain enough visible shape fill pixels.");
            Assert.True(textPixels > 30, "Rotated shape-text PNG baseline does not contain enough visible rotated text pixels.");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Rotated label", svg, StringComparison.Ordinal);
            Assert.Contains("transform=\"rotate(24", svg, StringComparison.Ordinal);
            Assert.Contains("#DBEAFE", svg, StringComparison.Ordinal);
            Assert.Contains("#2563EB", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void AlignedShapeTextImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateAlignedShapeTextBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:F8");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("Bottom right", svgText, StringComparison.Ordinal);
            Assert.Contains("text-anchor=\"end\"", svgText, StringComparison.Ordinal);
            Assert.Contains("#DCFCE7", svgText, StringComparison.Ordinal);
            Assert.Contains("#16A34A", svgText, StringComparison.Ordinal);
            AssertRasterBaseline(AlignedShapeTextBaselineName + ".png", png.Bytes);
            AssertTextBaseline(AlignedShapeTextBaselineName + ".svg", svgText);
        }

        [Fact]
        public void ApprovedAlignedShapeTextBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, AlignedShapeTextBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, AlignedShapeTextBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateAlignedShapeTextBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:F8");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(AlignedShapeTextBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(AlignedShapeTextBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved aligned shape-text PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved aligned shape-text SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved aligned shape-text PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 600, "Aligned shape-text PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 300, "Aligned shape-text PNG baseline height is unexpectedly small.");
            int fillPixels = CountPixelsNear(image, OfficeColor.FromRgb(220, 252, 231));
            int textPixels = CountPixelsNear(image, OfficeColor.FromRgb(31, 41, 55));
            Assert.True(fillPixels > 1200, "Aligned shape-text PNG baseline does not contain enough visible shape fill pixels.");
            Assert.True(textPixels > 30, "Aligned shape-text PNG baseline does not contain enough visible text pixels.");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Bottom right", svg, StringComparison.Ordinal);
            Assert.Contains("text-anchor=\"end\"", svg, StringComparison.Ordinal);
            Assert.Contains("#DCFCE7", svg, StringComparison.Ordinal);
            Assert.Contains("#16A34A", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void VerticalShapeTextImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateVerticalShapeTextBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:F8");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextVerticalOrientationUnsupported);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextVerticalOrientationUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("<text", svgText, StringComparison.Ordinal);
            Assert.DoesNotContain(">STACKED</text>", svgText, StringComparison.Ordinal);
            Assert.Contains(">S</text>", svgText, StringComparison.Ordinal);
            Assert.Contains(">T</text>", svgText, StringComparison.Ordinal);
            Assert.Contains(">A</text>", svgText, StringComparison.Ordinal);
            Assert.Contains(">C</text>", svgText, StringComparison.Ordinal);
            Assert.Contains(">K</text>", svgText, StringComparison.Ordinal);
            Assert.Contains(">E</text>", svgText, StringComparison.Ordinal);
            Assert.Contains(">D</text>", svgText, StringComparison.Ordinal);
            Assert.Contains("#E0F2FE", svgText, StringComparison.Ordinal);
            Assert.Contains("#0284C7", svgText, StringComparison.Ordinal);
            AssertRasterBaseline(VerticalShapeTextBaselineName + ".png", png.Bytes);
            AssertTextBaseline(VerticalShapeTextBaselineName + ".svg", svgText);
        }

        [Fact]
        public void ApprovedVerticalShapeTextBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, VerticalShapeTextBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, VerticalShapeTextBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateVerticalShapeTextBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:F8");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(VerticalShapeTextBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(VerticalShapeTextBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved vertical shape-text PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved vertical shape-text SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved vertical shape-text PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 600, "Vertical shape-text PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 300, "Vertical shape-text PNG baseline height is unexpectedly small.");
            int fillPixels = CountPixelsNear(image, OfficeColor.FromRgb(224, 242, 254));
            int textPixels = CountPixelsNear(image, OfficeColor.FromRgb(31, 41, 55));
            Assert.True(fillPixels > 1200, "Vertical shape-text PNG baseline does not contain enough visible shape fill pixels.");
            Assert.True(textPixels > 30, "Vertical shape-text PNG baseline does not contain enough visible stacked text pixels.");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("<text", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(">STACKED</text>", svg, StringComparison.Ordinal);
            Assert.Contains(">S</text>", svg, StringComparison.Ordinal);
            Assert.Contains(">T</text>", svg, StringComparison.Ordinal);
            Assert.Contains(">A</text>", svg, StringComparison.Ordinal);
            Assert.Contains(">C</text>", svg, StringComparison.Ordinal);
            Assert.Contains(">K</text>", svg, StringComparison.Ordinal);
            Assert.Contains(">E</text>", svg, StringComparison.Ordinal);
            Assert.Contains(">D</text>", svg, StringComparison.Ordinal);
            Assert.Contains("#E0F2FE", svg, StringComparison.Ordinal);
            Assert.Contains("#0284C7", svg, StringComparison.Ordinal);
        }

        private static ExcelBaselineFixture CreateRotatedPresetDrawingObjectBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelRotatedPresetDrawingObjectBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("RotatedShape");

            sheet.CellValue(1, 1, "Rotated DrawingML preset");
            sheet.Range("A1:F1").Merge();
            sheet.Range("A1:F1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);
            sheet.CellValue(7, 2, "Preset geometry, rotation, fill, outline, and clipping route through OfficeIMO.Drawing.");
            sheet.Range("B7:E7").Merge();
            sheet.Range("B7:E7").SetFillColor("F8FAFC").SetFontColor("334155");
            sheet.CellAlign(7, 2, HorizontalAlignmentValues.Center);

            for (int column = 1; column <= 6; column++) {
                sheet.SetColumnWidth(column, column == 1 || column == 6 ? 9 : 15);
            }

            sheet.SetRowHeight(1, 26);
            sheet.SetRowHeight(2, 42);
            sheet.SetRowHeight(3, 42);
            sheet.SetRowHeight(4, 42);
            sheet.SetRowHeight(5, 32);
            sheet.SetRowHeight(6, 28);
            sheet.SetRowHeight(7, 30);
            sheet.SetRowHeight(8, 22);
            for (int row = 1; row <= 8; row++) {
                for (int column = 1; column <= 6; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            AddRotatedPresetDrawingObjectShape(sheet);
            return new ExcelBaselineFixture(document, sheet);
        }

        private static ExcelBaselineFixture CreateRotatedShapeTextBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelRotatedShapeTextBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("RotatedText");

            sheet.CellValue(1, 1, "Rotated DrawingML text");
            sheet.Range("A1:F1").Merge();
            sheet.Range("A1:F1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);
            sheet.CellValue(7, 2, "Shape text rotates through OfficeIMO.Drawing and keeps an approximation diagnostic.");
            sheet.Range("B7:E7").Merge();
            sheet.Range("B7:E7").SetFillColor("F8FAFC").SetFontColor("334155");
            sheet.CellAlign(7, 2, HorizontalAlignmentValues.Center);

            for (int column = 1; column <= 6; column++) {
                sheet.SetColumnWidth(column, column == 1 || column == 6 ? 9 : 15);
            }

            sheet.SetRowHeight(1, 26);
            sheet.SetRowHeight(2, 42);
            sheet.SetRowHeight(3, 42);
            sheet.SetRowHeight(4, 42);
            sheet.SetRowHeight(5, 32);
            sheet.SetRowHeight(6, 28);
            sheet.SetRowHeight(7, 30);
            sheet.SetRowHeight(8, 22);
            for (int row = 1; row <= 8; row++) {
                for (int column = 1; column <= 6; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            AddRotatedShapeTextObject(sheet);
            return new ExcelBaselineFixture(document, sheet);
        }

        private static ExcelBaselineFixture CreateAlignedShapeTextBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelAlignedShapeTextBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("AlignedText");

            sheet.CellValue(1, 1, "Aligned DrawingML text");
            sheet.Range("A1:F1").Merge();
            sheet.Range("A1:F1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);
            sheet.CellValue(7, 2, "Shape text uses authored paragraph and body alignment through OfficeIMO.Drawing.");
            sheet.Range("B7:E7").Merge();
            sheet.Range("B7:E7").SetFillColor("F8FAFC").SetFontColor("334155");
            sheet.CellAlign(7, 2, HorizontalAlignmentValues.Center);

            for (int column = 1; column <= 6; column++) {
                sheet.SetColumnWidth(column, column == 1 || column == 6 ? 9 : 15);
            }

            sheet.SetRowHeight(1, 26);
            sheet.SetRowHeight(2, 42);
            sheet.SetRowHeight(3, 42);
            sheet.SetRowHeight(4, 42);
            sheet.SetRowHeight(5, 32);
            sheet.SetRowHeight(6, 28);
            sheet.SetRowHeight(7, 30);
            sheet.SetRowHeight(8, 22);
            for (int row = 1; row <= 8; row++) {
                for (int column = 1; column <= 6; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            AddAlignedShapeTextObject(sheet);
            return new ExcelBaselineFixture(document, sheet);
        }

        private static ExcelBaselineFixture CreateVerticalShapeTextBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelVerticalShapeTextBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("VerticalText");

            sheet.CellValue(1, 1, "Vertical DrawingML text");
            sheet.Range("A1:F1").Merge();
            sheet.Range("A1:F1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);
            sheet.CellValue(7, 2, "Simple vertical shape text routes through shared stacked text layout.");
            sheet.Range("B7:E7").Merge();
            sheet.Range("B7:E7").SetFillColor("F8FAFC").SetFontColor("334155");
            sheet.CellAlign(7, 2, HorizontalAlignmentValues.Center);

            for (int column = 1; column <= 6; column++) {
                sheet.SetColumnWidth(column, column == 1 || column == 6 ? 9 : 15);
            }

            sheet.SetRowHeight(1, 26);
            sheet.SetRowHeight(2, 42);
            sheet.SetRowHeight(3, 42);
            sheet.SetRowHeight(4, 42);
            sheet.SetRowHeight(5, 42);
            sheet.SetRowHeight(6, 22);
            sheet.SetRowHeight(7, 30);
            sheet.SetRowHeight(8, 22);
            for (int row = 1; row <= 8; row++) {
                for (int column = 1; column <= 6; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            AddVerticalShapeTextObject(sheet);
            return new ExcelBaselineFixture(document, sheet);
        }

        private static void AddRotatedPresetDrawingObjectShape(ExcelSheet sheet) {
            WorksheetPart worksheetPart = sheet.WorksheetPart;
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            var transform = new A.Transform2D {
                Rotation = (int)Math.Round(28D * 60000D)
            };

            drawingsPart.WorksheetDrawing.Append(new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("2"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("2"),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("4"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("5"),
                    new Xdr.RowOffset("0")),
                new Xdr.Shape(
                    new Xdr.NonVisualShapeProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 120U, Name = "Rotated heart" },
                        new Xdr.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true })),
                    new Xdr.ShapeProperties(
                        transform,
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Heart },
                        new A.SolidFill(new A.RgbColorModelHex { Val = "FDBA74" }),
                        new A.Outline(
                            new A.SolidFill(new A.RgbColorModelHex { Val = "EA580C" })) {
                            Width = 19050
                        }),
                    new Xdr.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph())),
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void AddRotatedShapeTextObject(ExcelSheet sheet) {
            WorksheetPart worksheetPart = sheet.WorksheetPart;
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            var transform = new A.Transform2D {
                Rotation = (int)Math.Round(24D * 60000D)
            };

            drawingsPart.WorksheetDrawing.Append(new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("2"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("2"),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("4"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("5"),
                    new Xdr.RowOffset("0")),
                new Xdr.Shape(
                    new Xdr.NonVisualShapeProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 121U, Name = "Rotated text box" },
                        new Xdr.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true })),
                    new Xdr.ShapeProperties(
                        transform,
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.RoundRectangle },
                        new A.SolidFill(new A.RgbColorModelHex { Val = "DBEAFE" }),
                        new A.Outline(
                            new A.SolidFill(new A.RgbColorModelHex { Val = "2563EB" })) {
                            Width = 19050
                        }),
                    new Xdr.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text("Rotated label"))))),
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void AddAlignedShapeTextObject(ExcelSheet sheet) {
            WorksheetPart worksheetPart = sheet.WorksheetPart;
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            drawingsPart.WorksheetDrawing.Append(new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("2"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("2"),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("4"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("5"),
                    new Xdr.RowOffset("0")),
                new Xdr.Shape(
                    new Xdr.NonVisualShapeProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 122U, Name = "Aligned text box" },
                        new Xdr.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true })),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(),
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.RoundRectangle },
                        new A.SolidFill(new A.RgbColorModelHex { Val = "DCFCE7" }),
                        new A.Outline(
                            new A.SolidFill(new A.RgbColorModelHex { Val = "16A34A" })) {
                            Width = 19050
                        }),
                    new Xdr.TextBody(
                        new A.BodyProperties { Anchor = A.TextAnchoringTypeValues.Bottom },
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.ParagraphProperties { Alignment = A.TextAlignmentTypeValues.Right },
                            new A.Run(new A.Text("Bottom right"))))),
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void AddVerticalShapeTextObject(ExcelSheet sheet) {
            WorksheetPart worksheetPart = sheet.WorksheetPart;
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            drawingsPart.WorksheetDrawing.Append(new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("2"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("2"),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("4"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("6"),
                    new Xdr.RowOffset("0")),
                new Xdr.Shape(
                    new Xdr.NonVisualShapeProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 123U, Name = "Vertical text box" },
                        new Xdr.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true })),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(),
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.RoundRectangle },
                        new A.SolidFill(new A.RgbColorModelHex { Val = "E0F2FE" }),
                        new A.Outline(
                            new A.SolidFill(new A.RgbColorModelHex { Val = "0284C7" })) {
                            Width = 19050
                        }),
                    new Xdr.TextBody(
                        new A.BodyProperties {
                            Anchor = A.TextAnchoringTypeValues.Center,
                            Vertical = A.TextVerticalValues.Vertical
                        },
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.ParagraphProperties { Alignment = A.TextAlignmentTypeValues.Center },
                            new A.Run(
                                new A.RunProperties { FontSize = 1400 },
                                new A.Text("STACKED"))))),
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

    }
}
