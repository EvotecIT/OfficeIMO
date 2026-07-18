using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using System.Globalization;
using System.Runtime.InteropServices;
using A = DocumentFormat.OpenXml.Drawing;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xunit;

namespace OfficeIMO.Tests {
    [Trait("Category", "ExcelImageVisualGate")]
    public partial class ExcelImageExportVisualBaselineTests {
        private const string BaselineName = "officeimo-excel-image-premium-range";
        private const string ConditionalBaselineName = "officeimo-excel-image-conditional-formatting";
        private const string ExpandedIconSetBaselineName = "officeimo-excel-image-expanded-icon-sets";
        private const string SparklineBaselineName = "officeimo-excel-image-sparklines";
        private const string ImageClippingBaselineName = "officeimo-excel-image-clipped-image";
        private const string TwoCellImageBaselineName = "officeimo-excel-image-two-cell-image";
        private const string CroppedImageBaselineName = "officeimo-excel-image-cropped-image";
        private const string RotatedImageBaselineName = "officeimo-excel-image-rotated-image";
        private const string TransformedImageBaselineName = "officeimo-excel-image-transformed-image";
        private const string DrawingObjectBaselineName = "officeimo-excel-image-drawing-object";
        private const string CommentBodyBaselineName = "officeimo-excel-image-comment-body";
        private const string RichTextBaselineName = "officeimo-excel-image-rich-text";
        private const string StackedTextBaselineName = "officeimo-excel-image-stacked-text";
        private const string PatternFillBaselineName = "officeimo-excel-image-pattern-fills";

        [Fact]
        public void PremiumRangeImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreatePremiumBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:H8");
            ExcelImageExportOptions options = CreateBaselineOptions();
            options.ShowCommentBodies = true;

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);

            OfficeImageExportDiagnostic commentDiagnostic = Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellCommentBodyApproximation);
            Assert.Equal("Premium!D7", commentDiagnostic.Source);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellCommentUnsupported);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            AssertDiagnosticsBaseline(BaselineName + ".diagnostics.txt", png.Diagnostics);
            AssertRasterBaseline(BaselineName + ".png", png.Bytes);
            AssertTextBaseline(BaselineName + ".svg", System.Text.Encoding.UTF8.GetString(svg.Bytes));
        }

        [Fact]
        public void ConditionalFormattingImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateConditionalFormattingBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:G7");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);

            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ConditionalIconSetApproximation);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Info, diagnostic.Severity);
            Assert.Equal("Signals!F3:F7", diagnostic.Source);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("#63B3ED", System.Text.Encoding.UTF8.GetString(svg.Bytes), StringComparison.Ordinal);
            Assert.Contains("#7C3AED", System.Text.Encoding.UTF8.GetString(svg.Bytes), StringComparison.Ordinal);
            Assert.Contains("#16A34A", System.Text.Encoding.UTF8.GetString(svg.Bytes), StringComparison.Ordinal);
            AssertDiagnosticsBaseline(ConditionalBaselineName + ".diagnostics.txt", png.Diagnostics);
            AssertRasterBaseline(ConditionalBaselineName + ".png", png.Bytes);
            AssertTextBaseline(ConditionalBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(svg.Bytes));
        }

        [Fact]
        public void ExpandedIconSetImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateExpandedIconSetBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:G7");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.Equal(4, png.Diagnostics.Count(item => item.Code == ExcelImageExportDiagnosticCodes.ConditionalIconSetApproximation));
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ConditionalIconSetUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("#16A34A", svgText, StringComparison.Ordinal);
            Assert.Contains("#F59E0B", svgText, StringComparison.Ordinal);
            Assert.Contains("#F97316", svgText, StringComparison.Ordinal);
            AssertDiagnosticsBaseline(ExpandedIconSetBaselineName + ".diagnostics.txt", png.Diagnostics);
            AssertRasterBaseline(ExpandedIconSetBaselineName + ".png", png.Bytes);
            AssertTextBaseline(ExpandedIconSetBaselineName + ".svg", svgText);
        }

        [Fact]
        public void SparklineImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateSparklineBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:E4");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);

            Assert.Equal(3, png.Diagnostics.Count(item => item.Code == ExcelImageExportDiagnosticCodes.SparklineRenderingApproximation));
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("<polyline", svgText, StringComparison.Ordinal);
            Assert.Contains("<circle", svgText, StringComparison.Ordinal);
            Assert.Contains("<rect", svgText, StringComparison.Ordinal);
            Assert.Contains("#2563EB", svgText, StringComparison.Ordinal);
            Assert.Contains("#16A34A", svgText, StringComparison.Ordinal);
            Assert.Contains("#DC2626", svgText, StringComparison.Ordinal);
            AssertDiagnosticsBaseline(SparklineBaselineName + ".diagnostics.txt", png.Diagnostics);
            AssertRasterBaseline(SparklineBaselineName + ".png", png.Bytes);
            AssertTextBaseline(SparklineBaselineName + ".svg", svgText);
        }

        [Fact]
        public void ImageClippingExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateImageClippingBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("B1:C3");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellTextOccludedByDrawing && item.Source == "ImageClip!B2");
            Assert.Contains(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellTextOccludedByDrawing && item.Source == "ImageClip!B2");
            Assert.Contains("clip-path=\"url(#xl-image-clip-", svgText, StringComparison.Ordinal);
            Assert.Contains("x=\"-", svgText, StringComparison.Ordinal);
            AssertDiagnosticsBaseline(ImageClippingBaselineName + ".diagnostics.txt", png.Diagnostics);
            AssertRasterBaseline(ImageClippingBaselineName + ".png", png.Bytes);
            AssertTextBaseline(ImageClippingBaselineName + ".svg", svgText);
        }

        [Fact]
        public void TwoCellImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateTwoCellImageBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:F6");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);
            AssertRasterBaseline(TwoCellImageBaselineName + ".png", png.Bytes);
            AssertTextBaseline(TwoCellImageBaselineName + ".svg", svgText);
        }

        [Fact]
        public void CroppedImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateCroppedImageBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:E5");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("clip-path=\"url(#xl-image-clip-", svgText, StringComparison.Ordinal);
            Assert.Contains("x=\"-", svgText, StringComparison.Ordinal);
            AssertRasterBaseline(CroppedImageBaselineName + ".png", png.Bytes);
            AssertTextBaseline(CroppedImageBaselineName + ".svg", svgText);
        }

        [Fact]
        public void RotatedImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateRotatedImageBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:E9");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("transform=\"rotate(30", svgText, StringComparison.Ordinal);
            AssertRasterBaseline(RotatedImageBaselineName + ".png", png.Bytes);
            AssertTextBaseline(RotatedImageBaselineName + ".svg", svgText);
        }

        [Fact]
        public void TransformedImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateTransformedImageBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:E9");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == "ExcelImageFlipUnsupported");
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == "ExcelImageCropRotationCombinationUnsupported");
            Assert.Contains("rotate(30", svgText, StringComparison.Ordinal);
            Assert.Contains("scale(-1 1)", svgText, StringComparison.Ordinal);
            AssertRasterBaseline(TransformedImageBaselineName + ".png", png.Bytes);
            AssertTextBaseline(TransformedImageBaselineName + ".svg", svgText);
        }

        [Fact]
        public void DrawingObjectImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateDrawingObjectBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:F6");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("Premium shape", svgText, StringComparison.Ordinal);
            Assert.Contains("#E0F2FE", svgText, StringComparison.Ordinal);
            Assert.Contains("#0284C7", svgText, StringComparison.Ordinal);
            AssertRasterBaseline(DrawingObjectBaselineName + ".png", png.Bytes);
            AssertTextBaseline(DrawingObjectBaselineName + ".svg", svgText);
        }

        [Fact]
        public void CommentBodyImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateCommentBodyBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:G7");
            ExcelImageExportOptions options = CreateBaselineOptions();
            options.ShowCommentBodies = true;

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellCommentBodyApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellCommentUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellTextOccludedByDrawing);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("Ready for leadership review", svgText, StringComparison.Ordinal);
            Assert.Contains("#FFFBE6", svgText, StringComparison.Ordinal);
            Assert.Contains("#FFF2CC", svgText, StringComparison.Ordinal);
            Assert.Contains("#C00000", svgText, StringComparison.Ordinal);
            AssertDiagnosticsBaseline(CommentBodyBaselineName + ".diagnostics.txt", png.Diagnostics);
            AssertRasterBaseline(CommentBodyBaselineName + ".png", png.Bytes);
            AssertTextBaseline(CommentBodyBaselineName + ".svg", svgText);
        }

        [Fact]
        public void RichTextImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateRichTextBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:B8");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.Contains(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && item.Source == "RichText!B6");
            Assert.Contains(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellTextRotationApproximation && item.Source == "RichText!B7");
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && item.Source == "RichText!B8");
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("Single", svgText, StringComparison.Ordinal);
            Assert.Contains("Hard", svgText, StringComparison.Ordinal);
            Assert.Contains("Wrapped", svgText, StringComparison.Ordinal);
            Assert.Contains("Shrink", svgText, StringComparison.Ordinal);
            Assert.Contains("Clip", svgText, StringComparison.Ordinal);
            Assert.Contains("Tilt", svgText, StringComparison.Ordinal);
            Assert.Contains("tiny line stays visible", svgText, StringComparison.Ordinal);
            Assert.Contains("transform=\"rotate(-45", svgText, StringComparison.Ordinal);
            Assert.Contains("#0F766E", svgText, StringComparison.Ordinal);
            Assert.Contains("#7C3AED", svgText, StringComparison.Ordinal);
            Assert.Contains("#DC2626", svgText, StringComparison.Ordinal);
            Assert.Contains("font-weight=\"700\"", svgText, StringComparison.Ordinal);
            Assert.Contains("font-style=\"italic\"", svgText, StringComparison.Ordinal);
            Assert.Contains("text-decoration=\"underline\"", svgText, StringComparison.Ordinal);
            AssertDiagnosticsBaseline(RichTextBaselineName + ".diagnostics.txt", png.Diagnostics);
            AssertRasterBaseline(RichTextBaselineName + ".png", png.Bytes);
            AssertTextBaseline(RichTextBaselineName + ".svg", svgText);
        }

        [Fact]
        public void StackedTextImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateStackedTextBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:D5");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.Equal(3, png.Diagnostics.Count(item => item.Code == ExcelImageExportDiagnosticCodes.CellTextRotationApproximation));
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellStackedTextRotationUnsupported);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellStackedTextRotationUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains(">S</text>", svgText, StringComparison.Ordinal);
            Assert.Contains(">K</text>", svgText, StringComparison.Ordinal);
            Assert.Contains(">R</text>", svgText, StringComparison.Ordinal);
            Assert.Contains("font-weight=\"700\"", svgText, StringComparison.Ordinal);
            Assert.Contains("font-style=\"italic\"", svgText, StringComparison.Ordinal);
            Assert.Contains("text-decoration=\"underline\"", svgText, StringComparison.Ordinal);
            Assert.DoesNotContain("rotate(", svgText, StringComparison.Ordinal);
            AssertDiagnosticsBaseline(StackedTextBaselineName + ".diagnostics.txt", png.Diagnostics);
            AssertRasterBaseline(StackedTextBaselineName + ".png", png.Bytes);
            AssertTextBaseline(StackedTextBaselineName + ".svg", svgText);
        }

        [Fact]
        public void PatternFillImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreatePatternFillBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:D5");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.Equal(7, png.Diagnostics.Count(item => item.Code == ExcelImageExportDiagnosticCodes.FillPatternApproximation));
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("stroke=\"#C00000\"", svgText, StringComparison.Ordinal);
            Assert.Contains("stroke=\"#1F4E79\"", svgText, StringComparison.Ordinal);
            Assert.Contains("fill=\"#70AD47\"", svgText, StringComparison.Ordinal);
            AssertDiagnosticsBaseline(PatternFillBaselineName + ".diagnostics.txt", png.Diagnostics);
            AssertRasterBaseline(PatternFillBaselineName + ".png", png.Bytes);
            AssertTextBaseline(PatternFillBaselineName + ".svg", svgText);
        }

        [Fact]
        public void ApprovedPremiumRangeBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, BaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, BaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreatePremiumBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:H8");
                ExcelImageExportOptions options = CreateBaselineOptions();
                options.ShowCommentBodies = true;
                AssertRasterBaseline(BaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(BaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved Excel PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved Excel SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved Excel PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 700, "Excel PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 300, "Excel PNG baseline height is unexpectedly small.");
            int nonBackgroundPixels = VisualBaselineTestSupport.CountNonBackgroundPixels(image, OfficeColor.White);
            int minimumVisiblePixels = Math.Max(500, image.Width * image.Height / 80);
            Assert.True(
                nonBackgroundPixels >= minimumVisiblePixels,
                "Excel PNG baseline appears blank or nearly blank. Visible pixels: " + nonBackgroundPixels + "/" + (image.Width * image.Height) + ".");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Operations Snapshot", svg, StringComparison.Ordinal);
            Assert.Contains("Revenue Trend", svg, StringComparison.Ordinal);
            Assert.Contains("94%", svg, StringComparison.Ordinal);
            Assert.Contains("71%", svg, StringComparison.Ordinal);
            Assert.Contains("82%", svg, StringComparison.Ordinal);
            Assert.Contains("Rich", svg, StringComparison.Ordinal);
            Assert.Contains(" text", svg, StringComparison.Ordinal);
            Assert.Contains("#0F766E", svg, StringComparison.Ordinal);
            Assert.Contains("#7C3AED", svg, StringComparison.Ordinal);
            Assert.Contains("font-weight=\"700\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-style=\"italic\"", svg, StringComparison.Ordinal);
            Assert.Contains("text-decoration=\"underline\"", svg, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svg, StringComparison.Ordinal);
            Assert.Contains("<clipPath", svg, StringComparison.Ordinal);
            Assert.Contains("<polygon", svg, StringComparison.Ordinal);
            Assert.Contains("#C00000", svg, StringComparison.Ordinal);
            Assert.Contains("Reviewer", svg, StringComparison.Ordinal);
            Assert.Contains("Ready for leadership review", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ApprovedRichTextBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, RichTextBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, RichTextBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateRichTextBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:B7");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(RichTextBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(RichTextBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved rich-text PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved rich-text SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved rich-text PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 350, "Rich-text PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 250, "Rich-text PNG baseline height is unexpectedly small.");
            int nonBackgroundPixels = VisualBaselineTestSupport.CountNonBackgroundPixels(image, OfficeColor.White);
            Assert.True(nonBackgroundPixels >= 900, "Rich-text PNG baseline appears blank or nearly blank. Visible pixels: " + nonBackgroundPixels + ".");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Rich Text Fidelity", svg, StringComparison.Ordinal);
            Assert.Contains("Single", svg, StringComparison.Ordinal);
            Assert.Contains("Hard", svg, StringComparison.Ordinal);
            Assert.Contains("Wrapped", svg, StringComparison.Ordinal);
            Assert.Contains("Shrink", svg, StringComparison.Ordinal);
            Assert.Contains("Clip", svg, StringComparison.Ordinal);
            Assert.Contains("Tilt", svg, StringComparison.Ordinal);
            Assert.Contains("transform=\"rotate(-45", svg, StringComparison.Ordinal);
            Assert.Contains("#0F766E", svg, StringComparison.Ordinal);
            Assert.Contains("#7C3AED", svg, StringComparison.Ordinal);
            Assert.Contains("#DC2626", svg, StringComparison.Ordinal);
            Assert.Contains("font-size", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ApprovedStackedTextBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, StackedTextBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, StackedTextBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateStackedTextBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:D5");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(StackedTextBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(StackedTextBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved stacked-text PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved stacked-text SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved stacked-text PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 420, "Stacked-text PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 230, "Stacked-text PNG baseline height is unexpectedly small.");
            int nonBackgroundPixels = VisualBaselineTestSupport.CountNonBackgroundPixels(image, OfficeColor.White);
            Assert.True(nonBackgroundPixels >= 1200, "Stacked-text PNG baseline appears blank or nearly blank. Visible pixels: " + nonBackgroundPixels + ".");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Stacked Text Fidelity", svg, StringComparison.Ordinal);
            Assert.Contains(">S</text>", svg, StringComparison.Ordinal);
            Assert.Contains(">K</text>", svg, StringComparison.Ordinal);
            Assert.Contains(">R</text>", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("rotate(", svg, StringComparison.Ordinal);
            Assert.Contains("#0F766E", svg, StringComparison.Ordinal);
            Assert.Contains("#7C3AED", svg, StringComparison.Ordinal);
            Assert.Contains("#DC2626", svg, StringComparison.Ordinal);
            Assert.Contains("#2563EB", svg, StringComparison.Ordinal);
            Assert.Contains("font-weight=\"700\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-style=\"italic\"", svg, StringComparison.Ordinal);
            Assert.Contains("text-decoration=\"underline\"", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ApprovedPatternFillBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, PatternFillBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, PatternFillBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreatePatternFillBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:D5");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(PatternFillBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(PatternFillBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved pattern-fill PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved pattern-fill SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved pattern-fill PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 400, "Pattern-fill PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 200, "Pattern-fill PNG baseline height is unexpectedly small.");
            int nonBackgroundPixels = VisualBaselineTestSupport.CountNonBackgroundPixels(image, OfficeColor.White);
            Assert.True(nonBackgroundPixels >= 1500, "Pattern-fill PNG baseline appears blank or nearly blank. Visible pixels: " + nonBackgroundPixels + ".");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Pattern Fill Fidelity", svg, StringComparison.Ordinal);
            Assert.Contains("Horizontal", svg, StringComparison.Ordinal);
            Assert.Contains("Grid", svg, StringComparison.Ordinal);
            Assert.Contains("Trellis", svg, StringComparison.Ordinal);
            Assert.Contains("stroke=\"#C00000\"", svg, StringComparison.Ordinal);
            Assert.Contains("stroke=\"#1F4E79\"", svg, StringComparison.Ordinal);
            Assert.Contains("fill=\"#70AD47\"", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ApprovedConditionalFormattingBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, ConditionalBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, ConditionalBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateConditionalFormattingBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:G7");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(ConditionalBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(ConditionalBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved conditional-formatting PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved conditional-formatting SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved conditional-formatting PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 800, "Conditional-formatting PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 250, "Conditional-formatting PNG baseline height is unexpectedly small.");
            int nonBackgroundPixels = VisualBaselineTestSupport.CountNonBackgroundPixels(image, OfficeColor.White);
            int minimumVisiblePixels = Math.Max(500, image.Width * image.Height / 90);
            Assert.True(
                nonBackgroundPixels >= minimumVisiblePixels,
                "Conditional-formatting PNG baseline appears blank or nearly blank. Visible pixels: " + nonBackgroundPixels + "/" + (image.Width * image.Height) + ".");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Conditional Signals", svg, StringComparison.Ordinal);
            Assert.Contains("Heat", svg, StringComparison.Ordinal);
            Assert.Contains("Load", svg, StringComparison.Ordinal);
            Assert.Contains("Rule", svg, StringComparison.Ordinal);
            Assert.Contains("#63B3ED", svg, StringComparison.Ordinal);
            Assert.Contains("#7C3AED", svg, StringComparison.Ordinal);
            Assert.Contains("#C6EFCE", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ApprovedExpandedIconSetBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, ExpandedIconSetBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, ExpandedIconSetBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateExpandedIconSetBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:G7");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(ExpandedIconSetBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(ExpandedIconSetBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved expanded-icon-set PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved expanded-icon-set SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved expanded-icon-set PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 600, "Expanded-icon-set PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 250, "Expanded-icon-set PNG baseline height is unexpectedly small.");
            int greenPixels = CountPixelsNear(image, OfficeColor.FromRgb(22, 163, 74));
            int orangePixels = CountPixelsNear(image, OfficeColor.FromRgb(249, 115, 22));
            Assert.True(greenPixels > 20, "Expanded-icon-set PNG baseline does not contain enough green icon pixels.");
            Assert.True(orangePixels > 20, "Expanded-icon-set PNG baseline does not contain enough orange icon pixels.");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Expanded Icon Sets", svg, StringComparison.Ordinal);
            Assert.Contains("Five arrows", svg, StringComparison.Ordinal);
            Assert.Contains("Four traffic", svg, StringComparison.Ordinal);
            Assert.Contains("#F59E0B", svg, StringComparison.Ordinal);
            Assert.Contains("#F97316", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ApprovedDrawingObjectBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, DrawingObjectBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, DrawingObjectBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateDrawingObjectBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:F6");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(DrawingObjectBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(DrawingObjectBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved drawing-object PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved drawing-object SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved drawing-object PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 600, "Drawing-object PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 250, "Drawing-object PNG baseline height is unexpectedly small.");
            int fillPixels = CountPixelsNear(image, OfficeColor.FromRgb(224, 242, 254));
            Assert.True(fillPixels > 1000, "Drawing-object PNG baseline does not contain enough visible shape fill pixels.");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Premium shape", svg, StringComparison.Ordinal);
            Assert.Contains("#E0F2FE", svg, StringComparison.Ordinal);
            Assert.Contains("#0284C7", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ApprovedCommentBodyBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, CommentBodyBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, CommentBodyBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateCommentBodyBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:G7");
                ExcelImageExportOptions options = CreateBaselineOptions();
                options.ShowCommentBodies = true;
                AssertRasterBaseline(CommentBodyBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(CommentBodyBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved comment-body PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved comment-body SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved comment-body PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 700, "Comment-body PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 250, "Comment-body PNG baseline height is unexpectedly small.");
            Assert.True(CountPixelsNear(image, OfficeColor.FromRgb(255, 251, 230)) > 500, "Comment-body PNG baseline does not contain enough callout fill pixels.");
            Assert.True(CountPixelsNear(image, OfficeColor.FromRgb(255, 242, 204)) > 100, "Comment-body PNG baseline does not contain enough callout header pixels.");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Comment Body Fidelity", svg, StringComparison.Ordinal);
            Assert.Contains("Ready for leadership review", svg, StringComparison.Ordinal);
            Assert.Contains("#FFFBE6", svg, StringComparison.Ordinal);
            Assert.Contains("#FFF2CC", svg, StringComparison.Ordinal);
            Assert.Contains("<polygon", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ApprovedSparklineBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, SparklineBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, SparklineBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateSparklineBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:E4");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(SparklineBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(SparklineBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved sparkline PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved sparkline SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved sparkline PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 700, "Sparkline PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 220, "Sparkline PNG baseline height is unexpectedly small.");
            int nonBackgroundPixels = VisualBaselineTestSupport.CountNonBackgroundPixels(image, OfficeColor.White);
            int minimumVisiblePixels = Math.Max(500, image.Width * image.Height / 100);
            Assert.True(
                nonBackgroundPixels >= minimumVisiblePixels,
                "Sparkline PNG baseline appears blank or nearly blank. Visible pixels: " + nonBackgroundPixels + "/" + (image.Width * image.Height) + ".");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Trend", svg, StringComparison.Ordinal);
            Assert.Contains("<polyline", svg, StringComparison.Ordinal);
            Assert.Contains("<circle", svg, StringComparison.Ordinal);
            Assert.Contains("<rect", svg, StringComparison.Ordinal);
            Assert.Contains("#2563EB", svg, StringComparison.Ordinal);
            Assert.Contains("#16A34A", svg, StringComparison.Ordinal);
            Assert.Contains("#DC2626", svg, StringComparison.Ordinal);
            Assert.Contains("clip-path=\"url(#officeimo-sparkline-clip-", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ApprovedImageClippingBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, ImageClippingBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, ImageClippingBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateImageClippingBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("B1:C3");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(ImageClippingBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(ImageClippingBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved clipped-image PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved clipped-image SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved clipped-image PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 250, "Clipped-image PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 160, "Clipped-image PNG baseline height is unexpectedly small.");
            int nonBackgroundPixels = VisualBaselineTestSupport.CountNonBackgroundPixels(image, OfficeColor.White);
            Assert.True(nonBackgroundPixels >= 500, "Clipped-image PNG baseline appears blank or nearly blank. Visible pixels: " + nonBackgroundPixels + ".");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("clip-path=\"url(#xl-image-clip-", svg, StringComparison.Ordinal);
            Assert.Contains("x=\"-", svg, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ApprovedTwoCellImageBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, TwoCellImageBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, TwoCellImageBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateTwoCellImageBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:F6");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(TwoCellImageBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(TwoCellImageBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved two-cell image PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved two-cell image SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved two-cell image PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 600, "Two-cell image PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 250, "Two-cell image PNG baseline height is unexpectedly small.");
            int nonBackgroundPixels = VisualBaselineTestSupport.CountNonBackgroundPixels(image, OfficeColor.White);
            Assert.True(nonBackgroundPixels >= 1200, "Two-cell image PNG baseline appears blank or nearly blank. Visible pixels: " + nonBackgroundPixels + ".");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Two-cell image anchor", svg, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ApprovedCroppedImageBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, CroppedImageBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, CroppedImageBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateCroppedImageBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:E5");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(CroppedImageBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(CroppedImageBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved cropped-image PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved cropped-image SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved cropped-image PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 700, "Cropped-image PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 250, "Cropped-image PNG baseline height is unexpectedly small.");
            int nonBackgroundPixels = VisualBaselineTestSupport.CountNonBackgroundPixels(image, OfficeColor.White);
            Assert.True(nonBackgroundPixels >= 1200, "Cropped-image PNG baseline appears blank or nearly blank. Visible pixels: " + nonBackgroundPixels + ".");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Cropped worksheet image", svg, StringComparison.Ordinal);
            Assert.Contains("x=\"-", svg, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ApprovedRotatedImageBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, RotatedImageBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, RotatedImageBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateRotatedImageBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:E9");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(RotatedImageBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(RotatedImageBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved rotated-image PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved rotated-image SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved rotated-image PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 700, "Rotated-image PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 250, "Rotated-image PNG baseline height is unexpectedly small.");
            int nonBackgroundPixels = VisualBaselineTestSupport.CountNonBackgroundPixels(image, OfficeColor.White);
            Assert.True(nonBackgroundPixels >= 1200, "Rotated-image PNG baseline appears blank or nearly blank. Visible pixels: " + nonBackgroundPixels + ".");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Rotated worksheet image", svg, StringComparison.Ordinal);
            Assert.Contains("transform=\"rotate(30", svg, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ApprovedTransformedImageBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, TransformedImageBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, TransformedImageBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateTransformedImageBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:E9");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(TransformedImageBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(TransformedImageBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved transformed-image PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved transformed-image SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved transformed-image PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 700, "Transformed-image PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 450, "Transformed-image PNG baseline height is unexpectedly small.");
            int nonBackgroundPixels = VisualBaselineTestSupport.CountNonBackgroundPixels(image, OfficeColor.White);
            Assert.True(nonBackgroundPixels >= 1200, "Transformed-image PNG baseline appears blank or nearly blank. Visible pixels: " + nonBackgroundPixels + ".");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("Cropped, flipped, rotated worksheet image", svg, StringComparison.Ordinal);
            Assert.Contains("rotate(30", svg, StringComparison.Ordinal);
            Assert.Contains("scale(-1 1)", svg, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", svg, StringComparison.Ordinal);
        }

        private static ExcelBaselineFixture CreatePremiumBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelImageBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Premium");

            sheet.CellValue(1, 1, "Operations Snapshot");
            sheet.Range("A1:H1").Merge();
            sheet.Range("A1:H1").SetFillColor("1F4E79").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);
            sheet.SetRowHeight(1, 26);

            sheet.CellValue(2, 1, "Region");
            sheet.CellValue(2, 2, "Score");
            sheet.CellValue(2, 3, "Status");
            sheet.CellValue(2, 4, "Narrative");
            sheet.Range("A2:D2").SetFillColor("D9EAF7").SetFontColor("1F2937").SetBold();

            sheet.CellValue(3, 1, "North");
            sheet.CellValue(3, 2, 0.94);
            sheet.CellAt(3, 2).Percent(0);
            sheet.CellValue(3, 3, "Ready");
            sheet.CellAt(3, 3).Success();
            sheet.CellValue(3, 4, "Wrapped cell text stays centered and readable");
            sheet.WrapCells(3, 3, 4);
            sheet.CellVerticalAlign(3, 4, VerticalAlignmentValues.Center);
            sheet.CellAlign(3, 4, HorizontalAlignmentValues.Center);

            sheet.CellValue(4, 1, "West");
            sheet.CellValue(4, 2, 0.71);
            sheet.CellAt(4, 2).Percent(0);
            sheet.CellValue(4, 3, "Watch");
            sheet.CellAt(4, 3).Warning();
            sheet.CellValue(4, 4, "Narrative stays bounded beside the chart");
            sheet.CellAt(4, 4).MutedText();
            sheet.WrapCells(4, 4, 4);
            sheet.CellVerticalAlign(4, 4, VerticalAlignmentValues.Center);

            sheet.CellValue(5, 1, "South");
            sheet.CellValue(5, 2, 0.82);
            sheet.CellAt(5, 2).Percent(0);
            sheet.CellValue(5, 3, "Risk");
            sheet.CellAt(5, 3).Error();
            sheet.CellValue(5, 4, "Bottom aligned note");
            sheet.CellVerticalAlign(5, 4, VerticalAlignmentValues.Bottom);
            sheet.CellAlign(5, 4, HorizontalAlignmentValues.Right);
            sheet.CellAt(6, 4).SetRichText(
                new ExcelRichTextRun("Rich") { Bold = true, FontColor = "0F766E", FontSize = 13D },
                new ExcelRichTextRun(" text") { Italic = true, Underline = true, FontColor = "7C3AED", FontSize = 12D });
            sheet.CellAt(6, 4).SetBorder(BorderStyleValues.Thin, "CBD5E1");
            sheet.CellAt(6, 4).SetFillColor("F8FAFC");
            sheet.CellValue(7, 4, "Review note");
            sheet.CellAt(7, 4).SetBorder(BorderStyleValues.Thin, "CBD5E1");
            sheet.CellAt(7, 4).SetFillColor("FFF7ED");
            sheet.CellAt(7, 4).SetFontColor("92400E");
            sheet.CellAlign(7, 4, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(7, 4, VerticalAlignmentValues.Center);
            sheet.SetComment("D7", "Ready for leadership review", "Reviewer");

            sheet.SetColumnWidth(1, 14);
            sheet.SetColumnWidth(2, 10);
            sheet.SetColumnWidth(3, 12);
            sheet.SetColumnWidth(4, 18);
            sheet.SetColumnWidth(5, 4);
            sheet.SetColumnWidth(6, 16);
            sheet.SetColumnWidth(7, 16);
            sheet.SetColumnWidth(8, 16);
            sheet.SetRowHeight(3, 54);
            sheet.SetRowHeight(4, 36);
            sheet.SetRowHeight(5, 40);
            sheet.SetRowHeight(6, 32);
            sheet.SetRowHeight(7, 32);

            for (int row = 2; row <= 5; row++) {
                for (int column = 1; column <= 4; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                }
            }

            sheet.AddImage(6, 2, CreateMarkerPng(), "image/png", widthPixels: 36, heightPixels: 22, name: "QualityMarker");
            sheet.AddChartFromRange("A2:B5", row: 2, column: 6, widthPixels: 230, heightPixels: 135, type: ExcelChartType.ColumnClustered, title: "Revenue Trend");
            return new ExcelBaselineFixture(document, sheet);
        }

        private static ExcelBaselineFixture CreateCommentBodyBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelCommentBodyBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("CommentBodies");

            sheet.CellValue(1, 1, "Comment Body Fidelity");
            sheet.Range("A1:G1").Merge();
            sheet.Range("A1:G1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);

            sheet.CellValue(2, 1, "Region");
            sheet.CellValue(2, 2, "Status");
            sheet.CellValue(2, 3, "Owner");
            sheet.CellValue(2, 4, "Note");
            sheet.Range("A2:D2").SetFillColor("E2E8F0").SetFontColor("0F172A").SetBold();

            sheet.CellValue(3, 1, "North");
            sheet.CellValue(3, 2, "Ready");
            sheet.CellAt(3, 2).Success();
            sheet.CellValue(3, 3, "Reviewer");
            sheet.CellValue(3, 4, "Open");
            sheet.SetComment("D3", "Ready for leadership review.\nCallout stays readable without hiding the table.", "Reviewer");

            sheet.CellValue(4, 1, "West");
            sheet.CellValue(4, 2, "Watch");
            sheet.CellAt(4, 2).Warning();
            sheet.CellValue(4, 3, "Planner");

            sheet.CellValue(5, 1, "South");
            sheet.CellValue(5, 2, "Risk");
            sheet.CellAt(5, 2).Error();
            sheet.CellValue(5, 3, "Owner");

            sheet.SetColumnWidth(1, 14);
            sheet.SetColumnWidth(2, 13);
            sheet.SetColumnWidth(3, 16);
            sheet.SetColumnWidth(4, 10);
            sheet.SetColumnWidth(5, 18);
            sheet.SetColumnWidth(6, 18);
            sheet.SetColumnWidth(7, 18);
            sheet.SetRowHeight(1, 28);
            for (int row = 2; row <= 7; row++) {
                sheet.SetRowHeight(row, row == 3 ? 32 : 28);
                for (int column = 1; column <= 7; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            sheet.Range("A3:D5").SetFillColor("F8FAFC");
            sheet.Range("E2:G7").SetFillColor("FFFFFF");
            return new ExcelBaselineFixture(document, sheet);
        }

        private static ExcelBaselineFixture CreateRichTextBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelRichTextBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("RichText");

            sheet.CellValue(1, 1, "Rich Text Fidelity");
            sheet.Range("A1:B1").Merge();
            sheet.Range("A1:B1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);

            string[] labels = { "Single", "Hard break", "Wrapped", "Shrink", "Clipped", "Rotated", "Mixed size" };
            for (int i = 0; i < labels.Length; i++) {
                int row = i + 2;
                sheet.CellValue(row, 1, labels[i]);
                sheet.CellAt(row, 1).SetFillColor("E2E8F0").SetFontColor("0F172A").SetBold();
                sheet.CellVerticalAlign(row, 1, VerticalAlignmentValues.Center);
                sheet.CellAt(row, 2).SetFillColor("F8FAFC").SetBorder(BorderStyleValues.Thin, "CBD5E1");
                sheet.CellVerticalAlign(row, 2, VerticalAlignmentValues.Center);
            }

            sheet.CellAt(2, 2).SetRichText(
                new ExcelRichTextRun("Single") { Bold = true, FontColor = "0F766E", FontSize = 13D },
                new ExcelRichTextRun(" styled") { Italic = true, FontColor = "7C3AED", FontSize = 12D },
                new ExcelRichTextRun(" line") { Underline = true, FontColor = "2563EB", FontSize = 12D });

            sheet.CellAt(3, 2).SetRichText(
                new ExcelRichTextRun("Hard") { Bold = true, FontColor = "DC2626", FontSize = 12D },
                new ExcelRichTextRun("\nbreak") { Italic = true, FontColor = "2563EB", FontSize = 12D });

            sheet.CellAt(4, 2).SetRichText(
                new ExcelRichTextRun("Wrapped ") { Bold = true, FontColor = "0F766E", FontSize = 12D },
                new ExcelRichTextRun("rich text keeps") { Italic = true, FontColor = "7C3AED", FontSize = 12D },
                new ExcelRichTextRun(" runs") { Underline = true, FontColor = "2563EB", FontSize = 12D });
            sheet.WrapCells(4, 4, 2);

            sheet.CellAt(5, 2)
                .SetShrinkToFit()
                .SetRichText(
                    new ExcelRichTextRun("Shrink") { Bold = true, FontColor = "0F766E", FontSize = 16D },
                    new ExcelRichTextRun(" to fit") { Italic = true, FontColor = "7C3AED", FontSize = 18D });

            sheet.CellAt(6, 2).SetRichText(
                new ExcelRichTextRun("Rich text stays bounded     hidden overflow") { Bold = true, FontColor = "DC2626", FontSize = 12D });

            sheet.CellAt(7, 2)
                .SetTextRotation(45)
                .SetRichText(
                    new ExcelRichTextRun("Tilt") { Bold = true, FontColor = "0F766E", FontSize = 14D },
                    new ExcelRichTextRun(" rich") { Italic = true, FontColor = "7C3AED", FontSize = 13D },
                    new ExcelRichTextRun(" text") { Underline = true, FontColor = "2563EB", FontSize = 13D });

            sheet.CellAt(8, 2).SetRichText(
                new ExcelRichTextRun("Large line") { Bold = true, FontColor = "0F766E", FontSize = 18D },
                new ExcelRichTextRun("\ntiny line stays visible") { Italic = true, FontColor = "7C3AED", FontSize = 8D });

            sheet.SetColumnWidth(1, 14);
            sheet.SetColumnWidth(2, 18);
            sheet.SetRowHeight(1, 28);
            sheet.SetRowHeight(2, 28);
            sheet.SetRowHeight(3, 46);
            sheet.SetRowHeight(4, 58);
            sheet.SetRowHeight(5, 28);
            sheet.SetRowHeight(6, 24);
            sheet.SetRowHeight(7, 48);
            sheet.SetRowHeight(8, 36);
            for (int row = 1; row <= 8; row++) {
                for (int column = 1; column <= 2; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                }
            }

            return new ExcelBaselineFixture(document, sheet);
        }

        private static ExcelBaselineFixture CreateStackedTextBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelStackedTextBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("StackedText");

            sheet.CellValue(1, 1, "Stacked Text Fidelity");
            sheet.Range("A1:D1").Merge();
            sheet.Range("A1:D1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);

            string[] headers = { "Case", "Status", "Narrow", "Marker" };
            for (int column = 1; column <= headers.Length; column++) {
                sheet.CellValue(2, column, headers[column - 1]);
                sheet.CellAt(2, column).SetFillColor("E2E8F0").SetFontColor("0F172A").SetBold();
                sheet.CellAlign(2, column, HorizontalAlignmentValues.Center);
                sheet.CellVerticalAlign(2, column, VerticalAlignmentValues.Center);
            }

            sheet.CellValue(3, 1, "Centered");
            sheet.CellValue(4, 1, "Shrink");
            sheet.CellValue(5, 1, "Mixed");
            sheet.Range("A3:A5").SetFillColor("F8FAFC").SetFontColor("334155");

            sheet.CellValue(3, 2, "STACK");
            sheet.CellAt(3, 2).SetTextRotation(255).SetFontColor("0F766E").SetBold().SetFontSize(12);
            sheet.CellAlign(3, 2, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(3, 2, VerticalAlignmentValues.Center);

            sheet.CellValue(3, 3, "EXPORT");
            sheet.CellAt(3, 3).SetTextRotation(255).SetFontColor("7C3AED").SetBold().SetShrinkToFit().SetFontSize(14);
            sheet.CellAlign(3, 3, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(3, 3, VerticalAlignmentValues.Center);

            sheet.CellAt(3, 4).SetTextRotation(255).SetRichText(
                new ExcelRichTextRun("R") { Bold = true, FontColor = "DC2626", FontSize = 12D },
                new ExcelRichTextRun("E") { Italic = true, FontColor = "EA580C", FontSize = 12D },
                new ExcelRichTextRun("A") { Underline = true, FontColor = "2563EB", FontSize = 12D },
                new ExcelRichTextRun("D") { Bold = true, FontColor = "16A34A", FontSize = 12D },
                new ExcelRichTextRun("Y") { FontColor = "7C3AED", FontSize = 12D });
            sheet.CellAlign(3, 4, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(3, 4, VerticalAlignmentValues.Center);

            sheet.CellValue(4, 2, "PNG");
            sheet.CellValue(4, 3, "SVG");
            sheet.CellValue(4, 4, "Drawing-owned stacked layout");
            sheet.CellAt(4, 2).SetFontColor("0F766E").SetBold();
            sheet.CellAt(4, 3).SetFontColor("7C3AED").SetBold();
            sheet.CellAt(4, 4).SetFontColor("475569").SetShrinkToFit();

            sheet.CellValue(5, 2, "Shared layout");
            sheet.CellValue(5, 3, "No old unsupported diagnostic");
            sheet.CellValue(5, 4, "PNG/SVG baseline gate");
            sheet.Range("B5:D5").SetFontColor("475569");
            sheet.WrapCells(5, 5, 3);

            sheet.SetColumnWidth(1, 14);
            sheet.SetColumnWidth(2, 12);
            sheet.SetColumnWidth(3, 10);
            sheet.SetColumnWidth(4, 20);
            sheet.SetRowHeight(1, 28);
            sheet.SetRowHeight(2, 26);
            sheet.SetRowHeight(3, 96);
            sheet.SetRowHeight(4, 30);
            sheet.SetRowHeight(5, 42);

            for (int row = 1; row <= 5; row++) {
                for (int column = 1; column <= 4; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            sheet.Range("B3:D3").SetFillColor("F8FAFC");
            sheet.Range("B4:D5").SetFillColor("FFFFFF");
            return new ExcelBaselineFixture(document, sheet);
        }

        private static ExcelBaselineFixture CreatePatternFillBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelPatternFillBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Patterns");

            sheet.CellValue(1, 1, "Pattern Fill Fidelity");
            sheet.Range("A1:D1").Merge();
            sheet.Range("A1:D1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);

            string[] headers = { "Axis", "Red", "Blue", "Green" };
            for (int column = 1; column <= headers.Length; column++) {
                sheet.CellValue(2, column, headers[column - 1]);
                sheet.CellAt(2, column).SetFillColor("E2E8F0").SetFontColor("0F172A").SetBold();
                sheet.CellAlign(2, column, HorizontalAlignmentValues.Center);
            }

            sheet.CellValue(3, 1, "Light");
            sheet.CellValue(3, 2, "Horizontal");
            sheet.CellValue(3, 3, "Vertical");
            sheet.CellValue(3, 4, "Grid");
            sheet.CellValue(4, 1, "Diagonal");
            sheet.CellValue(4, 2, "Down");
            sheet.CellValue(4, 3, "Up");
            sheet.CellValue(4, 4, "Trellis");
            sheet.CellValue(5, 1, "Gray");
            sheet.CellValue(5, 2, "Dots");
            sheet.CellValue(5, 3, "Fallback");
            sheet.CellValue(5, 4, "Approx");

            ApplyPatternFill(sheet, 3, 2, PatternValues.LightHorizontal, "FFC00000", "FFFFE5E5");
            ApplyPatternFill(sheet, 3, 3, PatternValues.LightVertical, "FF1F4E79", "FFDDEBF7");
            ApplyPatternFill(sheet, 3, 4, PatternValues.LightGrid, "FF70AD47", "FFE2F0D9");
            ApplyPatternFill(sheet, 4, 2, PatternValues.DarkDown, "FFC00000", "FFFFE5E5");
            ApplyPatternFill(sheet, 4, 3, PatternValues.DarkUp, "FF1F4E79", "FFDDEBF7");
            ApplyPatternFill(sheet, 4, 4, PatternValues.DarkTrellis, "FF70AD47", "FFE2F0D9");
            ApplyPatternFill(sheet, 5, 2, PatternValues.Gray125, "FF70AD47", "FFF8FAFC");

            sheet.SetColumnWidth(1, 13);
            sheet.SetColumnWidth(2, 15);
            sheet.SetColumnWidth(3, 15);
            sheet.SetColumnWidth(4, 15);
            sheet.SetRowHeight(1, 28);
            for (int row = 2; row <= 5; row++) {
                sheet.SetRowHeight(row, 32);
                for (int column = 1; column <= 4; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                    sheet.CellAlign(row, column, HorizontalAlignmentValues.Center);
                }
            }

            return new ExcelBaselineFixture(document, sheet);
        }

        private static ExcelBaselineFixture CreateConditionalFormattingBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelConditionalBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Signals");

            sheet.CellValue(1, 1, "Conditional Signals");
            sheet.Range("A1:G1").Merge();
            sheet.Range("A1:G1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);
            sheet.SetRowHeight(1, 26);

            string[] headers = { "Service", "Heat", "Load", "Delta", "State", "Icons", "Rule" };
            for (int column = 1; column <= headers.Length; column++) {
                sheet.CellValue(2, column, headers[column - 1]);
                sheet.CellAt(2, column).SetFillColor("E2E8F0").SetFontColor("1F2937").SetBold();
            }

            string[] services = { "Auth", "Mail", "Files", "Sync", "Edge" };
            int[] heat = { 12, 41, 63, 78, 94 };
            int[] load = { 18, 39, 57, 74, 96 };
            int[] delta = { -18, -6, 4, 16, 28 };
            int[] ruleValues = { 5, 20, 30, 8, 16 };
            string[] states = { "Quiet", "Normal", "Busy", "Hot", "Critical" };
            for (int i = 0; i < services.Length; i++) {
                int row = i + 3;
                sheet.CellValue(row, 1, services[i]);
                sheet.CellValue(row, 2, heat[i]);
                sheet.CellValue(row, 3, load[i]);
                sheet.CellValue(row, 4, delta[i]);
                sheet.CellValue(row, 5, states[i]);
                sheet.CellValue(row, 6, i + 1);
                sheet.CellValue(row, 7, ruleValues[i]);
                sheet.CellAt(row, 5).SetFillColor(i >= 3 ? "FEE2E2" : "F8FAFC");
            }

            sheet.SetColumnWidth(1, 14);
            sheet.SetColumnWidth(2, 11);
            sheet.SetColumnWidth(3, 16);
            sheet.SetColumnWidth(4, 16);
            sheet.SetColumnWidth(5, 13);
            sheet.SetColumnWidth(6, 10);
            sheet.SetColumnWidth(7, 11);
            for (int row = 2; row <= 7; row++) {
                sheet.SetRowHeight(row, 24);
                for (int column = 1; column <= 7; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            sheet.CellAlign(2, 2, HorizontalAlignmentValues.Center);
            sheet.CellAlign(2, 3, HorizontalAlignmentValues.Center);
            sheet.CellAlign(2, 4, HorizontalAlignmentValues.Center);
            sheet.CellAlign(2, 6, HorizontalAlignmentValues.Center);
            sheet.CellAlign(2, 7, HorizontalAlignmentValues.Center);
            for (int row = 3; row <= 7; row++) {
                sheet.CellAlign(row, 2, HorizontalAlignmentValues.Center);
                sheet.CellAlign(row, 3, HorizontalAlignmentValues.Center);
                sheet.CellAlign(row, 4, HorizontalAlignmentValues.Center);
                sheet.CellAlign(row, 6, HorizontalAlignmentValues.Center);
                sheet.CellAlign(row, 7, HorizontalAlignmentValues.Center);
            }

            sheet.AddConditionalColorScale("B3:B7", OfficeColor.FromRgb(254, 202, 202), OfficeColor.FromRgb(34, 197, 94));
            sheet.AddConditionalDataBar("C3:C7", OfficeColor.FromRgb(99, 179, 237));
            sheet.AddConditionalDataBar("D3:D7", OfficeColor.FromRgb(124, 58, 237));
            sheet.AddConditionalIconSet("F3:F7");
            sheet.AddConditionalRule("G3:G7", ConditionalFormattingOperatorValues.GreaterThan, "15", fillColor: "C6EFCE");
            return new ExcelBaselineFixture(document, sheet);
        }

        private static ExcelBaselineFixture CreateExpandedIconSetBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelExpandedIconSetBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("IconSets");

            sheet.CellValue(1, 1, "Expanded Icon Sets");
            sheet.Range("A1:G1").Merge();
            sheet.Range("A1:E1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);
            sheet.SetRowHeight(1, 26);

            string[] headers = { "Tier", "Five arrows", "Arrow value", "Four traffic", "Traffic value", "Five rating", "Five quarters" };
            for (int column = 1; column <= headers.Length; column++) {
                sheet.CellValue(2, column, headers[column - 1]);
                sheet.CellAt(2, column).SetFillColor("E2E8F0").SetFontColor("1F2937").SetBold();
                sheet.CellVerticalAlign(2, column, VerticalAlignmentValues.Center);
            }

            string[] tiers = { "Lowest", "Low", "Middle", "High", "Highest" };
            for (int i = 0; i < tiers.Length; i++) {
                int row = i + 3;
                sheet.CellValue(row, 1, tiers[i]);
                sheet.CellValue(row, 2, i + 1);
                sheet.CellValue(row, 3, i + 1);
                sheet.CellValue(row, 4, Math.Min(i + 1, 4));
                sheet.CellValue(row, 5, Math.Min(i + 1, 4));
                sheet.CellValue(row, 6, i + 1);
                sheet.CellValue(row, 7, i + 1);
            }

            for (int column = 1; column <= 7; column++) {
                sheet.SetColumnWidth(column, column == 1 ? 13 : 12);
            }

            for (int row = 2; row <= 7; row++) {
                sheet.SetRowHeight(row, 24);
                for (int column = 1; column <= 7; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                    sheet.CellAlign(row, column, column == 1 ? HorizontalAlignmentValues.Left : HorizontalAlignmentValues.Center);
                }
            }

            sheet.AddConditionalIconSet("B3:B7", IconSetValues.FiveArrows, showValue: true, reverseIconOrder: false);
            sheet.AddConditionalIconSet("D3:D7", IconSetValues.FourTrafficLights, showValue: true, reverseIconOrder: false);
            sheet.AddConditionalIconSet("F3:F7", IconSetValues.FiveRating, showValue: true, reverseIconOrder: false);
            sheet.AddConditionalIconSet("G3:G7", IconSetValues.FiveQuarters, showValue: true, reverseIconOrder: false);
            return new ExcelBaselineFixture(document, sheet);
        }

        private static ExcelBaselineFixture CreateSparklineBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelSparklineBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Sparklines");

            sheet.CellValue(1, 1, "Metric");
            sheet.CellValue(1, 2, "Jan");
            sheet.CellValue(1, 3, "Feb");
            sheet.CellValue(1, 4, "Mar");
            sheet.CellValue(1, 5, "Trend");
            sheet.Range("A1:E1").SetFillColor("EAF2F8").SetFontColor("0F172A").SetBold();
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 2, VerticalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 3, VerticalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 4, VerticalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 5, VerticalAlignmentValues.Center);

            sheet.CellValue(2, 1, "Revenue");
            sheet.CellValue(2, 2, 10);
            sheet.CellValue(2, 3, 18);
            sheet.CellValue(2, 4, 14);
            sheet.CellValue(3, 1, "Margin");
            sheet.CellValue(3, 2, 8);
            sheet.CellValue(3, 3, -4);
            sheet.CellValue(3, 4, 12);
            sheet.CellValue(4, 1, "Wins");
            sheet.CellValue(4, 2, 1);
            sheet.CellValue(4, 3, -1);
            sheet.CellValue(4, 4, 1);

            sheet.SetColumnWidth(1, 14);
            sheet.SetColumnWidth(2, 9);
            sheet.SetColumnWidth(3, 9);
            sheet.SetColumnWidth(4, 9);
            sheet.SetColumnWidth(5, 16);
            sheet.SetRowHeight(1, 28);
            sheet.SetRowHeight(2, 30);
            sheet.SetRowHeight(3, 30);
            sheet.SetRowHeight(4, 30);
            for (int row = 1; row <= 4; row++) {
                for (int column = 1; column <= 5; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "E2E8F0");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            for (int row = 2; row <= 4; row++) {
                sheet.CellAlign(row, 2, HorizontalAlignmentValues.Center);
                sheet.CellAlign(row, 3, HorizontalAlignmentValues.Center);
                sheet.CellAlign(row, 4, HorizontalAlignmentValues.Center);
            }

            sheet.AddSparklines("B2:D2", "E2", displayMarkers: true, displayHighLow: true, displayAxis: true, seriesColor: "#2563EB", markersColor: "#1D4ED8", highColor: "#16A34A", lowColor: "#DC2626", axisColor: "#94A3B8");
            sheet.AddSparklines("B3:D3", "E3", SparklineTypeValues.Column, displayNegative: true, displayAxis: true, seriesColor: "#16A34A", negativeColor: "#DC2626", axisColor: "#94A3B8");
            sheet.AddSparklines("B4:D4", "E4", SparklineTypeValues.Stacked, displayNegative: true, displayAxis: true, seriesColor: "#0EA5E9", negativeColor: "#DC2626", axisColor: "#94A3B8");
            return new ExcelBaselineFixture(document, sheet);
        }

        private static ExcelBaselineFixture CreateImageClippingBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelImageClippingBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ImageClip");

            sheet.CellValue(1, 1, "Outside Anchor");
            sheet.CellValue(1, 2, "Selected Range");
            sheet.Range("B1:C1").Merge();
            sheet.Range("B1:C1").SetFillColor("EAF2F8").SetFontColor("0F172A").SetBold();
            sheet.CellAlign(1, 2, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 2, VerticalAlignmentValues.Center);
            sheet.CellValue(2, 2, "Only the overlapping slice is exported");
            sheet.CellValue(3, 2, "The image anchor remains in column A");
            sheet.Range("B2:C3").SetFillColor("F8FAFC");
            sheet.SetColumnWidth(1, 10);
            sheet.SetColumnWidth(2, 16);
            sheet.SetColumnWidth(3, 16);
            sheet.SetRowHeight(1, 26);
            sheet.SetRowHeight(2, 42);
            sheet.SetRowHeight(3, 42);
            for (int row = 1; row <= 3; row++) {
                for (int column = 1; column <= 3; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            sheet.AddImage(2, 1, CreateClippedBannerPng(), "image/png", widthPixels: 170, heightPixels: 44, name: "WideBanner");
            return new ExcelBaselineFixture(document, sheet);
        }

        private static ExcelBaselineFixture CreateTwoCellImageBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelTwoCellImageBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("TwoCell");

            sheet.CellValue(1, 1, "Two-cell image anchor");
            sheet.Range("A1:F1").Merge();
            sheet.Range("A1:F1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);
            sheet.CellValue(5, 2, "Picture spans B2:E5 from marker geometry");
            sheet.Range("B5:E5").Merge();
            sheet.Range("B5:E5").SetFillColor("F8FAFC").SetFontColor("334155");
            sheet.CellAlign(5, 2, HorizontalAlignmentValues.Center);

            for (int column = 1; column <= 6; column++) {
                sheet.SetColumnWidth(column, column == 1 || column == 6 ? 9 : 14);
            }

            sheet.SetRowHeight(1, 26);
            sheet.SetRowHeight(2, 30);
            sheet.SetRowHeight(3, 30);
            sheet.SetRowHeight(4, 30);
            sheet.SetRowHeight(5, 26);
            sheet.SetRowHeight(6, 22);
            for (int row = 1; row <= 6; row++) {
                for (int column = 1; column <= 6; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            AddTwoCellAnchoredImage(sheet, CreateTwoCellBannerPng());
            return new ExcelBaselineFixture(document, sheet);
        }

        private static ExcelBaselineFixture CreateCroppedImageBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelCroppedImageBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Crop");

            sheet.CellValue(1, 1, "Cropped worksheet image");
            sheet.Range("A1:E1").Merge();
            sheet.Range("A1:E1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);
            sheet.CellValue(5, 1, "Source has red, blue, and green bands; srcRect crops to the blue center.");
            sheet.Range("A5:E5").Merge();
            sheet.Range("A5:E5").SetFillColor("F8FAFC").SetFontColor("334155");
            sheet.CellAlign(5, 1, HorizontalAlignmentValues.Center);

            for (int column = 1; column <= 5; column++) {
                sheet.SetColumnWidth(column, 14);
            }

            sheet.SetRowHeight(1, 26);
            sheet.SetRowHeight(2, 34);
            sheet.SetRowHeight(3, 34);
            sheet.SetRowHeight(4, 34);
            sheet.SetRowHeight(5, 28);
            for (int row = 1; row <= 5; row++) {
                for (int column = 1; column <= 5; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            AddCroppedImage(sheet, CreateCroppedBandPng());
            return new ExcelBaselineFixture(document, sheet);
        }

        private static ExcelBaselineFixture CreateRotatedImageBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelRotatedImageBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Rotate");

            sheet.CellValue(1, 1, "Rotated worksheet image");
            sheet.Range("A1:E1").Merge();
            sheet.Range("A1:E1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);
            sheet.CellValue(9, 1, "The picture keeps its Excel transform and rotates around its center.");
            sheet.Range("A9:E9").Merge();
            sheet.Range("A9:E9").SetFillColor("F8FAFC").SetFontColor("334155");
            sheet.CellAlign(9, 1, HorizontalAlignmentValues.Center);

            for (int column = 1; column <= 5; column++) {
                sheet.SetColumnWidth(column, 14);
            }

            sheet.SetRowHeight(1, 26);
            sheet.SetRowHeight(2, 34);
            sheet.SetRowHeight(3, 34);
            sheet.SetRowHeight(4, 34);
            sheet.SetRowHeight(5, 34);
            sheet.SetRowHeight(6, 34);
            sheet.SetRowHeight(7, 34);
            sheet.SetRowHeight(8, 34);
            sheet.SetRowHeight(9, 28);
            for (int row = 1; row <= 9; row++) {
                for (int column = 1; column <= 5; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            AddRotatedImage(sheet, CreateRotatedBannerPng());
            return new ExcelBaselineFixture(document, sheet);
        }

        private static ExcelBaselineFixture CreateTransformedImageBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelTransformedImageBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Transform");

            sheet.CellValue(1, 1, "Cropped, flipped, rotated worksheet image");
            sheet.Range("A1:E1").Merge();
            sheet.Range("A1:E1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);
            sheet.CellValue(9, 1, "Source crop, horizontal flip, and rotation share one Drawing image projector.");
            sheet.Range("A9:E9").Merge();
            sheet.Range("A9:E9").SetFillColor("F8FAFC").SetFontColor("334155");
            sheet.CellAlign(9, 1, HorizontalAlignmentValues.Center);

            for (int column = 1; column <= 5; column++) {
                sheet.SetColumnWidth(column, 14);
            }

            sheet.SetRowHeight(1, 26);
            for (int row = 2; row <= 8; row++) {
                sheet.SetRowHeight(row, 34);
            }

            sheet.SetRowHeight(9, 28);
            for (int row = 1; row <= 9; row++) {
                for (int column = 1; column <= 5; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            AddTransformedImage(sheet, CreateTransformedBannerPng());
            return new ExcelBaselineFixture(document, sheet);
        }

        private static ExcelBaselineFixture CreateDrawingObjectBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelDrawingObjectBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Drawing");

            sheet.CellValue(1, 1, "Worksheet drawing object");
            sheet.Range("A1:F1").Merge();
            sheet.Range("A1:F1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);
            sheet.CellValue(5, 2, "Simple text shapes render through OfficeIMO.Drawing.");
            sheet.Range("B5:E5").Merge();
            sheet.Range("B5:E5").SetFillColor("F8FAFC").SetFontColor("334155");
            sheet.CellAlign(5, 2, HorizontalAlignmentValues.Center);

            for (int column = 1; column <= 6; column++) {
                sheet.SetColumnWidth(column, column == 1 || column == 6 ? 9 : 14);
            }

            sheet.SetRowHeight(1, 26);
            sheet.SetRowHeight(2, 38);
            sheet.SetRowHeight(3, 38);
            sheet.SetRowHeight(4, 30);
            sheet.SetRowHeight(5, 28);
            sheet.SetRowHeight(6, 22);
            for (int row = 1; row <= 6; row++) {
                for (int column = 1; column <= 6; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            AddDrawingObjectShape(sheet);
            return new ExcelBaselineFixture(document, sheet);
        }

        private static ExcelImageExportOptions CreateBaselineOptions() =>
            new ExcelImageExportOptions {
                Scale = 2,
                ShowGridlines = false,
                IncludeImages = true,
                IncludeCharts = true,
                BackgroundColor = OfficeColor.White
            };

        private static byte[] CreateMarkerPng() {
            OfficeRasterImage image = new OfficeRasterImage(36, 22, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            canvas.FillRectangle(1, 1, 34, 20, OfficeColor.FromRgb(220, 252, 231));
            canvas.DrawRectangle(1, 1, 34, 20, OfficeColor.FromRgb(22, 163, 74));
            canvas.DrawLine(8, 11, 15, 17, OfficeColor.FromRgb(22, 101, 52), 2);
            canvas.DrawLine(15, 17, 28, 5, OfficeColor.FromRgb(22, 101, 52), 2);
            return OfficePngWriter.Encode(image, OfficePngCompression.Stored);
        }

        private static byte[] CreateClippedBannerPng() {
            OfficeRasterImage image = new OfficeRasterImage(170, 44, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            canvas.FillRectangle(0, 0, 72, 44, OfficeColor.FromRgb(220, 38, 38));
            canvas.FillRectangle(72, 0, 98, 44, OfficeColor.FromRgb(37, 99, 235));
            canvas.DrawRectangle(1, 1, 168, 42, OfficeColor.FromRgb(15, 23, 42), 2);
            canvas.DrawLine(72, 2, 72, 42, OfficeColor.White, 2);
            canvas.DrawLine(86, 10, 150, 10, OfficeColor.White, 2);
            canvas.DrawLine(86, 22, 138, 22, OfficeColor.White, 2);
            canvas.DrawLine(86, 34, 124, 34, OfficeColor.White, 2);
            return OfficePngWriter.Encode(image, OfficePngCompression.Stored);
        }

        private static byte[] CreateTwoCellBannerPng() {
            OfficeRasterImage image = new OfficeRasterImage(240, 96, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            canvas.FillRectangle(0, 0, 240, 96, OfficeColor.FromRgb(37, 99, 235));
            canvas.FillRectangle(0, 0, 64, 96, OfficeColor.FromRgb(20, 184, 166));
            canvas.FillRectangle(176, 0, 64, 96, OfficeColor.FromRgb(124, 58, 237));
            canvas.DrawRectangle(2, 2, 236, 92, OfficeColor.White, 3);
            canvas.DrawLine(86, 28, 154, 28, OfficeColor.White, 4);
            canvas.DrawLine(64, 48, 176, 48, OfficeColor.White, 4);
            canvas.DrawLine(86, 68, 154, 68, OfficeColor.White, 4);
            return OfficePngWriter.Encode(image, OfficePngCompression.Stored);
        }

        private static byte[] CreateCroppedBandPng() {
            OfficeRasterImage image = new OfficeRasterImage(200, 80, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            canvas.FillRectangle(0, 0, 50, 80, OfficeColor.FromRgb(220, 38, 38));
            canvas.FillRectangle(50, 0, 100, 80, OfficeColor.FromRgb(37, 99, 235));
            canvas.FillRectangle(150, 0, 50, 80, OfficeColor.FromRgb(22, 163, 74));
            canvas.DrawRectangle(52, 6, 96, 68, OfficeColor.White, 3);
            canvas.DrawLine(76, 26, 124, 26, OfficeColor.White, 4);
            canvas.DrawLine(66, 42, 134, 42, OfficeColor.White, 4);
            canvas.DrawLine(76, 58, 124, 58, OfficeColor.White, 4);
            return OfficePngWriter.Encode(image, OfficePngCompression.Stored);
        }

        private static byte[] CreateRotatedBannerPng() {
            OfficeRasterImage image = new OfficeRasterImage(220, 84, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            canvas.FillRectangle(0, 0, 220, 84, OfficeColor.FromRgb(37, 99, 235));
            canvas.FillRectangle(0, 0, 64, 84, OfficeColor.FromRgb(20, 184, 166));
            canvas.FillRectangle(156, 0, 64, 84, OfficeColor.FromRgb(124, 58, 237));
            canvas.DrawRectangle(3, 3, 214, 78, OfficeColor.White, 4);
            canvas.DrawLine(82, 26, 138, 26, OfficeColor.White, 5);
            canvas.DrawLine(64, 43, 156, 43, OfficeColor.White, 5);
            canvas.DrawLine(82, 60, 138, 60, OfficeColor.White, 5);
            return OfficePngWriter.Encode(image, OfficePngCompression.Stored);
        }

        private static byte[] CreateTransformedBannerPng() {
            OfficeRasterImage image = new OfficeRasterImage(240, 90, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            canvas.FillRectangle(0, 0, 60, 90, OfficeColor.FromRgb(220, 38, 38));
            canvas.FillRectangle(60, 0, 112, 90, OfficeColor.FromRgb(37, 99, 235));
            canvas.FillRectangle(172, 0, 68, 90, OfficeColor.FromRgb(22, 163, 74));
            canvas.DrawRectangle(64, 8, 102, 74, OfficeColor.White, 4);
            canvas.DrawLine(92, 32, 142, 32, OfficeColor.White, 5);
            canvas.DrawLine(82, 52, 152, 52, OfficeColor.White, 5);
            return OfficePngWriter.Encode(image, OfficePngCompression.Stored);
        }

        private static void AddDrawingObjectShape(ExcelSheet sheet) {
            WorksheetPart worksheetPart = sheet.WorksheetPart;
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            drawingsPart.WorksheetDrawing.Append(new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("1"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("1"),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("4"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("3"),
                    new Xdr.RowOffset("0")),
                new Xdr.Shape(
                    new Xdr.NonVisualShapeProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 110U, Name = "Premium shape" },
                        new Xdr.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true })),
                    new Xdr.ShapeProperties(
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.RoundRectangle },
                        new A.SolidFill(new A.RgbColorModelHex { Val = "E0F2FE" }),
                        new A.Outline(
                            new A.SolidFill(new A.RgbColorModelHex { Val = "0284C7" })) {
                            Width = 12700
                        }),
                    new Xdr.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(new A.Run(new A.Text("Premium shape"))))),
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void AddTwoCellAnchoredImage(ExcelSheet sheet, byte[] imageBytes) {
            WorksheetPart worksheetPart = sheet.WorksheetPart;
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            ImagePart imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
            using (MemoryStream stream = new MemoryStream(imageBytes)) {
                imagePart.FeedData(stream);
            }

            string relationshipId = drawingsPart.GetIdOfPart(imagePart);
            drawingsPart.WorksheetDrawing.Append(new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("1"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("1"),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("5"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("4"),
                    new Xdr.RowOffset("0")),
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 93U, Name = "TwoCellBanner" },
                        new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = relationshipId },
                        new A.Stretch(new A.FillRectangle())),
                    new Xdr.ShapeProperties(
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void AddCroppedImage(ExcelSheet sheet, byte[] imageBytes) {
            WorksheetPart worksheetPart = sheet.WorksheetPart;
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            ImagePart imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
            using (MemoryStream stream = new MemoryStream(imageBytes)) {
                imagePart.FeedData(stream);
            }

            string relationshipId = drawingsPart.GetIdOfPart(imagePart);
            drawingsPart.WorksheetDrawing.Append(new Xdr.OneCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("0"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("1"),
                    new Xdr.RowOffset("0")),
                new Xdr.Extent { Cx = 520L * 9525L, Cy = 102L * 9525L },
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 94U, Name = "CroppedBand" },
                        new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = relationshipId },
                        new A.SourceRectangle { Left = 25000, Right = 25000 },
                        new A.Stretch(new A.FillRectangle())),
                    new Xdr.ShapeProperties(
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void AddRotatedImage(ExcelSheet sheet, byte[] imageBytes) {
            WorksheetPart worksheetPart = sheet.WorksheetPart;
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            ImagePart imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
            using (MemoryStream stream = new MemoryStream(imageBytes)) {
                imagePart.FeedData(stream);
            }

            string relationshipId = drawingsPart.GetIdOfPart(imagePart);
            drawingsPart.WorksheetDrawing.Append(new Xdr.OneCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("1"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("3"),
                    new Xdr.RowOffset("0")),
                new Xdr.Extent { Cx = 300L * 9525L, Cy = 96L * 9525L },
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 95U, Name = "RotatedBanner" },
                        new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = relationshipId },
                        new A.Stretch(new A.FillRectangle())),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0L, Y = 0L },
                            new A.Extents { Cx = 300L * 9525L, Cy = 96L * 9525L }) { Rotation = 30 * 60000 },
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void AddTransformedImage(ExcelSheet sheet, byte[] imageBytes) {
            WorksheetPart worksheetPart = sheet.WorksheetPart;
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            ImagePart imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
            using (MemoryStream stream = new MemoryStream(imageBytes)) {
                imagePart.FeedData(stream);
            }

            string relationshipId = drawingsPart.GetIdOfPart(imagePart);
            drawingsPart.WorksheetDrawing.Append(new Xdr.OneCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("1"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("3"),
                    new Xdr.RowOffset("0")),
                new Xdr.Extent { Cx = 300L * 9525L, Cy = 104L * 9525L },
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 96U, Name = "TransformedBanner" },
                        new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = relationshipId },
                        new A.SourceRectangle { Left = 25000 },
                        new A.Stretch(new A.FillRectangle())),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0L, Y = 0L },
                            new A.Extents { Cx = 300L * 9525L, Cy = 104L * 9525L }) {
                            Rotation = 30 * 60000,
                            HorizontalFlip = true
                        },
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void ApplyPatternFill(ExcelSheet sheet, int row, int column, PatternValues pattern, string foregroundArgb, string backgroundArgb) {
            WorkbookPart workbookPart = sheet.WorksheetPart.GetParentParts().OfType<WorkbookPart>().Single();
            WorkbookStylesPart stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            stylesheet.Fills ??= new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
            stylesheet.CellFormats ??= new CellFormats(new CellFormat());

            stylesheet.Fills.Append(new Fill(new PatternFill {
                PatternType = pattern,
                ForegroundColor = new ForegroundColor { Rgb = foregroundArgb },
                BackgroundColor = new BackgroundColor { Rgb = backgroundArgb }
            }));
            uint fillId = (uint)stylesheet.Fills.Count();
            stylesheet.Fills.Count = fillId;

            Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<Cell>()
                .Single(item => string.Equals(item.CellReference?.Value, A1.CellReference(row, column), StringComparison.OrdinalIgnoreCase));
            CellFormat baseFormat = stylesheet.CellFormats.Elements<CellFormat>().ElementAtOrDefault((int)(cell.StyleIndex?.Value ?? 0U)) ?? new CellFormat();
            CellFormat format = (CellFormat)baseFormat.CloneNode(true);
            format.FillId = fillId - 1U;
            format.ApplyFill = true;
            stylesheet.CellFormats.Append(format);
            uint styleIndex = (uint)stylesheet.CellFormats.Count();
            stylesheet.CellFormats.Count = styleIndex;
            cell.StyleIndex = styleIndex - 1U;
            stylesheet.Save();
        }

        private static int CountPixelsNear(OfficeRasterImage image, OfficeColor expected) {
            int count = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor color = image.GetPixel(x, y);
                    if (Math.Abs(color.R - expected.R) <= 8 &&
                        Math.Abs(color.G - expected.G) <= 8 &&
                        Math.Abs(color.B - expected.B) <= 8 &&
                        color.A >= 248) {
                        count++;
                    }
                }
            }

            return count;
        }

        private static void AssertRasterBaseline(string baselineName, byte[] actualPng) {
            string expectedPath = Path.Combine(BaselineDirectory, baselineName);
            if (UpdateBaselines) {
                Directory.CreateDirectory(Path.GetDirectoryName(expectedPath)!);
                File.WriteAllBytes(expectedPath, actualPng);
                return;
            }

            if (!File.Exists(expectedPath)) {
                throw new FileNotFoundException(
                    "Excel image baseline missing. Set OFFICEIMO_UPDATE_EXCEL_IMAGE_BASELINES=1 and re-run this test to generate it.",
                    expectedPath);
            }

            if (!ShouldCompareApprovedBaselinesExactly) {
                Assert.True(OfficePngReader.TryDecode(actualPng, out OfficeRasterImage? rendered), "Actual Excel image export is not a supported PNG file.");
                Assert.NotNull(rendered);
                Assert.True(rendered!.Width > 0 && rendered.Height > 0, "Actual Excel image export must have non-zero dimensions.");
                return;
            }

            int channelTolerance = VisualBaselineTestSupport.ReadNonNegativeInt("OFFICEIMO_EXCEL_IMAGE_BASELINE_PIXEL_TOLERANCE", 0);
            int allowedDifferentPixels = VisualBaselineTestSupport.ReadNonNegativeInt("OFFICEIMO_EXCEL_IMAGE_BASELINE_ALLOWED_DIFF_PIXELS", 0);
            double maximumMeanAbsoluteError = VisualBaselineTestSupport.ReadNonNegativeDouble("OFFICEIMO_EXCEL_IMAGE_BASELINE_MAX_MAE", 0D);
            double maximumRootMeanSquareError = VisualBaselineTestSupport.ReadNonNegativeDouble("OFFICEIMO_EXCEL_IMAGE_BASELINE_MAX_RMSE", 0D);
            double maximumMeanLuminanceError = VisualBaselineTestSupport.ReadNonNegativeDouble("OFFICEIMO_EXCEL_IMAGE_BASELINE_MAX_LUMINANCE_MAE", 0D);
            VisualRasterComparison comparison = VisualBaselineTestSupport.CompareRasterImages(
                File.ReadAllBytes(expectedPath),
                actualPng,
                channelTolerance,
                allowedDifferentPixels,
                maximumMeanAbsoluteError,
                maximumRootMeanSquareError,
                maximumMeanLuminanceError);
            if (comparison.Passed) {
                return;
            }

            string artifactDirectory = VisualBaselineTestSupport.CreateArtifactDirectory("OfficeIMO.ExcelImageBaselines");
            File.WriteAllBytes(Path.Combine(artifactDirectory, "actual-" + baselineName), actualPng);
            File.Copy(expectedPath, Path.Combine(artifactDirectory, "expected-" + baselineName), overwrite: true);
            File.WriteAllBytes(Path.Combine(artifactDirectory, Path.GetFileNameWithoutExtension(baselineName) + ".diff.png"), comparison.DiffPng);
            throw new Xunit.Sdk.XunitException(
                "Excel image raster baseline changed for '" + baselineName + "'. " +
                "Different pixels: " + comparison.DifferentPixels + "/" + comparison.TotalPixels + "; " +
                "max channel delta: " + comparison.MaxChannelDelta + "; " +
                "allowed different pixels: " + comparison.AllowedDifferentPixels + "; " +
                "channel tolerance: " + comparison.ChannelTolerance + "; " +
                "MAE: " + comparison.MeanAbsoluteError.ToString("0.###", CultureInfo.InvariantCulture) + "/" +
                    comparison.MaximumMeanAbsoluteError.ToString("0.###", CultureInfo.InvariantCulture) + "; " +
                "RMSE: " + comparison.RootMeanSquareError.ToString("0.###", CultureInfo.InvariantCulture) + "/" +
                    comparison.MaximumRootMeanSquareError.ToString("0.###", CultureInfo.InvariantCulture) + "; " +
                "luminance MAE: " + comparison.MeanLuminanceError.ToString("0.###", CultureInfo.InvariantCulture) + "/" +
                    comparison.MaximumMeanLuminanceError.ToString("0.###", CultureInfo.InvariantCulture) + ". " +
                "Artifacts: " + artifactDirectory + ".");
        }

        private static void AssertTextBaseline(string baselineName, string actualText) {
            string expectedPath = Path.Combine(BaselineDirectory, baselineName);
            string normalizedActual = VisualBaselineTestSupport.NormalizeText(actualText);
            if (UpdateBaselines) {
                Directory.CreateDirectory(Path.GetDirectoryName(expectedPath)!);
                File.WriteAllText(expectedPath, normalizedActual, new System.Text.UTF8Encoding(false));
                return;
            }

            if (!File.Exists(expectedPath)) {
                throw new FileNotFoundException(
                    "Excel image SVG baseline missing. Set OFFICEIMO_UPDATE_EXCEL_IMAGE_BASELINES=1 and re-run this test to generate it.",
                    expectedPath);
            }

            string expectedText = VisualBaselineTestSupport.NormalizeText(File.ReadAllText(expectedPath));
            if (string.Equals(expectedText, normalizedActual, StringComparison.Ordinal)) {
                return;
            }

            string artifactDirectory = VisualBaselineTestSupport.CreateArtifactDirectory("OfficeIMO.ExcelImageBaselines");
            File.WriteAllText(Path.Combine(artifactDirectory, "actual-" + baselineName), normalizedActual, new System.Text.UTF8Encoding(false));
            File.Copy(expectedPath, Path.Combine(artifactDirectory, "expected-" + baselineName), overwrite: true);
            throw new Xunit.Sdk.XunitException("Excel image SVG baseline changed for '" + baselineName + "'. Artifacts: " + artifactDirectory + ".");
        }

        private static void AssertDiagnosticsBaseline(string baselineName, IReadOnlyList<OfficeImageExportDiagnostic> diagnostics) {
            string expectedPath = Path.Combine(BaselineDirectory, baselineName);
            string normalizedActual = CreateDiagnosticsBaselineText(diagnostics);
            if (UpdateBaselines) {
                Directory.CreateDirectory(Path.GetDirectoryName(expectedPath)!);
                File.WriteAllText(expectedPath, normalizedActual, new System.Text.UTF8Encoding(false));
                return;
            }

            if (!File.Exists(expectedPath)) {
                throw new FileNotFoundException(
                    "Excel image diagnostics baseline missing. Set OFFICEIMO_UPDATE_EXCEL_IMAGE_BASELINES=1 and re-run this test to generate it.",
                    expectedPath);
            }

            string expectedText = VisualBaselineTestSupport.NormalizeText(File.ReadAllText(expectedPath));
            if (string.Equals(expectedText, normalizedActual, StringComparison.Ordinal)) {
                return;
            }

            if (!ShouldCompareApprovedBaselinesExactly) {
                Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
                return;
            }

            string artifactDirectory = VisualBaselineTestSupport.CreateArtifactDirectory("OfficeIMO.ExcelImageBaselines");
            File.WriteAllText(Path.Combine(artifactDirectory, "actual-" + baselineName), normalizedActual, new System.Text.UTF8Encoding(false));
            File.Copy(expectedPath, Path.Combine(artifactDirectory, "expected-" + baselineName), overwrite: true);
            throw new Xunit.Sdk.XunitException("Excel image diagnostics baseline changed for '" + baselineName + "'. Artifacts: " + artifactDirectory + ".");
        }

        private static string CreateDiagnosticsBaselineText(IReadOnlyList<OfficeImageExportDiagnostic> diagnostics) {
            var builder = new System.Text.StringBuilder();
            foreach (OfficeImageExportDiagnostic diagnostic in diagnostics
                         .OrderBy(item => item.Source ?? string.Empty, StringComparer.Ordinal)
                         .ThenBy(item => item.Code, StringComparer.Ordinal)
                         .ThenBy(item => item.Message, StringComparer.Ordinal)) {
                builder
                    .Append(diagnostic.Severity)
                    .Append('|')
                    .Append(diagnostic.Code)
                    .Append('|')
                    .Append(diagnostic.Source ?? string.Empty)
                    .Append('|')
                    .Append(VisualBaselineTestSupport.NormalizeText(diagnostic.Message).Replace("\n", "\\n"))
                    .Append('\n');
            }

            return builder.ToString();
        }

        private static bool ShouldCompareApprovedBaselinesExactly =>
            RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ||
            string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_EXCEL_IMAGE_STRICT_CROSS_PLATFORM_BASELINES"), "1", StringComparison.Ordinal);

        private static string BaselineDirectory =>
            Path.Combine(VisualBaselineTestSupport.GetTestsProjectRoot(), "Excel", "VisualBaselines");

        private static bool UpdateBaselines =>
            string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_EXCEL_IMAGE_BASELINES"), "1", StringComparison.Ordinal);

        private sealed class ExcelBaselineFixture : IDisposable {
            internal ExcelBaselineFixture(ExcelDocument document, ExcelSheet sheet) {
                Document = document;
                Sheet = sheet;
            }

            private ExcelDocument Document { get; }

            internal ExcelSheet Sheet { get; }

            public void Dispose() {
                Document.Dispose();
            }
        }
    }
}
