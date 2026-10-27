using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfAuthoringLayoutDiagnosticsTests {
    [Theory]
    [InlineData("flow", false)]
    [InlineData("column", false)]
    [InlineData("canvas", false)]
    [InlineData("text", false)]
    [InlineData("textbox", false)]
    [InlineData("flow", true)]
    [InlineData("column", true)]
    [InlineData("canvas", true)]
    [InlineData("text", true)]
    [InlineData("textbox", true)]
    public void RasterInlineReportsNonzeroPaintBeyondTheTextFrame(string frame, bool visible) {
        var report = new PdfConversionReport();
        byte[] png = PdfPngTestImages.CreateRgbaPng(255, 0, 0, visible ? (byte)128 : (byte)0);
        var image = new PdfInlineImage(png, 24, 40);
        var runs = new[] { PdfTextRun.Inline(image) };
        var options = CreateOptions(report);
        byte[] bytes;
        if (frame is "text" or "textbox") {
            bytes = PdfDocument.Create(options).Canvas(canvas => {
                if (frame == "text") canvas.Text(runs, 24, 24, 120, 18);
                else canvas.TextBox(runs, 24, 24, 120, 18, new PdfCanvasTextBoxStyle { FontSize = 9, PaddingX = 0, PaddingY = 0 });
            }).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        } else {
            // Exact leading makes image overhang reachable without admitting an oversized line.
            var cell = new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs,
                lineSpacing: PdfLineSpacing.Exactly(12), widowControl: false) });
            var style = CreateTableStyle(18);
            style.ClipTextToCellBounds = true;
            bytes = RenderPaintTable(frame, options, cell, style);
        }
        Assert.NotEmpty(bytes);
        string code = frame == "text" ? "canvas-text-paint-clipped" : frame == "textbox" ? "canvas-text-box-paint-clipped" : "table-cell-paint-clipped";
        Assert.Equal(visible, report.Warnings.Any(warning => warning.Code == code));
        Assert.DoesNotContain(report.Warnings, warning => warning.Code.EndsWith("PaintBoundsUnmeasured", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RasterInlineOmittedBeforeRowPaintStillReportsOnlyNonzeroAlpha(bool visible) {
        var report = new PdfConversionReport();
        byte[] png = PdfPngTestImages.CreateRgbaPng(255, 0, 0, visible ? (byte)128 : (byte)0);
        var cell = new PdfTableCell(new[] { PdfTextRun.Inline(new PdfInlineImage(png, 24, 40)) });
        byte[] bytes = RenderPaintTable("flow", CreateOptions(report), cell, CreateTableStyle(18));
        Assert.NotEmpty(bytes);
        Assert.Equal(visible, report.Warnings.Any(warning => warning.Code == "table-cell-inline-paint-omitted"));
    }

    [Theory]
    [InlineData(OfficeImageFit.Contain, 32D, false)]
    [InlineData(OfficeImageFit.Stretch, 32D, true)]
    [InlineData(OfficeImageFit.Cover, 32D, true)]
    [InlineData(OfficeImageFit.Cover, 48D, false)]
    public void RasterInlineUsesFittedPaintAndDoesNotCountItsOwnCoverCropAsLoss(OfficeImageFit fit, double height, bool clipped) {
        var report = new PdfConversionReport();
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(80, 8, OfficeColor.Red));
        byte[] bytes = PdfDocument.Create(CreateOptions(report)).Canvas(canvas => canvas.Text(new[] {
            PdfTextRun.Inline(new PdfInlineImage(png, 24, 40, fit: fit))
        }, 24, 24, 120, height)).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.Single(PdfImageExtractor.ExtractImagePlacements(bytes));
        Assert.Equal(clipped, report.Warnings.Any(warning => warning.Code == "canvas-text-paint-clipped"));
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "canvas-text-canvas-clip-paint-clipped");
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "canvas-text-paint-bounds-unmeasured");
    }

    [Theory]
    [InlineData(6, false)]
    [InlineData(8, true)]
    public void RasterInlineProjectsTheAlphaSupportThroughExifOrientation(int orientation, bool clipped) {
        var raster = new OfficeRasterImage(2, 1);
        raster.SetPixel(0, 0, OfficeColor.Red); // Right half stays transparent.
        var metadata = new OfficeImageMetadata();
        metadata.SetExifValue(OfficeExifTag.Orientation, (ushort)orientation);
        byte[] png = OfficeImageMetadata.Apply(OfficePngWriter.Encode(raster), metadata);
        var report = new PdfConversionReport();
        byte[] bytes = PdfDocument.Create(CreateOptions(report)).Canvas(canvas => canvas.Text(new[] {
            PdfTextRun.Inline(new PdfInlineImage(png, 40, 40, fit: OfficeImageFit.Stretch))
        }, 24, 24, 120, 24)).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.Single(PdfImageExtractor.ExtractImagePlacements(bytes));
        Assert.Equal(clipped, report.Warnings.Any(warning => warning.Code == "canvas-text-paint-clipped"));
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "canvas-text-paint-bounds-unmeasured");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RasterAlphaOutsideTheAncestorClipDoesNotCauseASecondFrameLoss(bool paintOutsideAncestor) {
        var raster = new OfficeRasterImage(4, 1);
        raster.SetPixel(paintOutsideAncestor ? 3 : 0, 0, OfficeColor.Red);
        var report = new PdfConversionReport();
        byte[] bytes = PdfDocument.Create(CreateOptions(report)).Canvas(canvas => canvas.Clip(24, 24, 10, 48,
            child => child.Text(new[] { PdfTextRun.Inline(new PdfInlineImage(OfficePngWriter.Encode(raster), 40, 40, fit: OfficeImageFit.Stretch)) },
                24, 24, 40, 48))).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.NotEmpty(bytes);
        Assert.Equal(paintOutsideAncestor, report.Warnings.Any(warning => warning.Code == "canvas-text-canvas-clip-paint-clipped"));
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "canvas-text-paint-clipped");
    }

    [Fact]
    public void RasterAlphaInspectionLimitReportsUnassessedWithoutRejectingTheDocument() {
        var report = new PdfConversionReport();
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1025, 1024, OfficeColor.FromRgba(255, 0, 0, 128)));
        byte[] bytes = PdfDocument.Create(CreateOptions(report)).Canvas(canvas => canvas.Text(new[] {
            PdfTextRun.Inline(new PdfInlineImage(png, 24, 40))
        }, 24, 24, 120, 18)).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.NotEmpty(bytes);
        var warning = Assert.Single(report.Warnings, warning => warning.Code == "canvas-text-paint-bounds-unmeasured");
        Assert.Equal(OfficeConversionLossKind.Unassessed, warning.LossKind);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "canvas-text-paint-clipped");
    }

    [Fact]
    public void NominalStandardGlyphBoundsDoNotProvePartialPathInkLoss() {
        var report = new PdfConversionReport();
        OfficeClipPath triangle = OfficeClipPath.Path(OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(8, 0),
            OfficePathCommand.LineTo(0, 12), OfficePathCommand.Close());
        byte[] bytes = PdfDocument.Create(CreateOptions(report)).Canvas(canvas => canvas.Clip(24, 24, triangle,
            child => child.Text(new[] { PdfTextRun.Normal("O", font: PdfStandardFont.TimesRoman) }, 24, 24, 12, 18))).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.NotEmpty(bytes);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "canvas-text-canvas-clip-paint-clipped");
        var warning = Assert.Single(report.Warnings, warning => warning.Code == "canvas-text-paint-bounds-unmeasured");
        Assert.Equal(OfficeConversionLossKind.Unassessed, warning.LossKind);
    }

    [Fact]
    public void MeasurementSnapshotBorrowsAttachmentsAndKeepsGlyphUsageIndependent() {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Typography", "OfficeIMOBaselineSans-Regular.ttf"));
        var source = new PdfCreateOptions().EmbedStandardFont(PdfStandardFont.Helvetica, font, "Snapshot font")
            .AddEmbeddedFile("attachment.bin", new byte[65536]);
        Assert.True(source.TryGetEmbeddedStandardFontProgramForGeneration(PdfStandardFont.Helvetica, out _, out var sourceFont));
        var measurement = source.CreateTextMeasurementSnapshot();
        Assert.Same(Assert.Single(source.EmbeddedFilesForInspection), Assert.Single(measurement.EmbeddedFilesForInspection));
        Assert.NotNull(PdfWriter.MeasurePositionedText(PdfTextRun.Normal("Z"), measurement));
        Assert.Empty(sourceFont!.GetUsedGlyphIds());
        Assert.True(measurement.TryGetEmbeddedStandardFontProgramForGeneration(PdfStandardFont.Helvetica, out _, out var measuredFont));
        Assert.NotEmpty(measuredFont!.GetUsedGlyphIds());
    }

    [Fact]
    public void RasterDiagnosticsWithAttachmentsDoNotChangeSerializedBytes() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(4, 1, OfficeColor.FromRgba(255, 0, 0, 128)));
        byte[] Render(bool diagnostics) {
            var report = new PdfConversionReport();
            var options = CreateOptions(report);
            if (!diagnostics) options = new PdfCreateOptions { PageWidth = 260, PageHeight = 180, MarginLeft = 24, MarginRight = 24,
                MarginTop = 24, MarginBottom = 24, DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 9 };
            options.AddEmbeddedFile("attachment.bin", new byte[65536]);
            return PdfDocument.Create(options).Canvas(canvas => canvas.Text(new[] {
                PdfTextRun.Inline(new PdfInlineImage(png, 40, 40))
            }, 24, 24, 120, 18)).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        }
        Assert.Equal(Render(false), Render(true));
    }
}
