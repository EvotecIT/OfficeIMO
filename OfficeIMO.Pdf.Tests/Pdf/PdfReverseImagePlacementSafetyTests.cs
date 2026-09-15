using OfficeIMO.Html.Pdf;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfReverseImagePlacementSafetyTests {
    private static readonly byte[] Png = PdfPngTestImages.CreateRgbPng(8, 4);

    [Fact]
    public void InvisibleImagePlacementsAreSuppressedAcrossEditableAdapters() {
        byte[] source = CreateDocument()
            .Canvas(canvas => canvas.Effect(
                OfficeIMO.Drawing.OfficeTransform.Identity,
                0D,
                effect => effect.Image(Png, 20D, 30D, 80D, 40D)))
            .ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);
        Assert.Equal(0D, Assert.Single(Assert.Single(logical.Pages).Images).Placements[0].Opacity);

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Empty(word.Value.Images);
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfInvisibleImagePlacementSuppressed" &&
                warning.LossKind == OfficeConversionLossKind.None);
        }

        PdfHtmlConversionResult html = logical.ToHtmlResult(PdfToHtmlOptions.CreateSemanticProfile());
        Assert.DoesNotContain("data:image/", html.Value, StringComparison.Ordinal);
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == "InvisibleImagePlacementSuppressed" &&
            warning.LossKind == OfficeConversionLossKind.None);

        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using (powerPoint.Value) {
            Assert.Empty(powerPoint.Value.Slides.SelectMany(static slide => slide.Pictures));
            Assert.Equal(0, Assert.Single(powerPoint.Report.EditablePages).OmittedImageCount);
            Assert.Contains(powerPoint.Report.Warnings, static warning =>
                warning.Code == "PdfInvisibleImagePlacementSuppressed" &&
                warning.LossKind == OfficeConversionLossKind.None);
        }
    }

    [Fact]
    public void InvisibleImagesDoNotCreateOmissionLossWhenImageOutputIsDisabled() {
        byte[] invisible = CreateDocument()
            .Canvas(canvas => canvas.Effect(
                OfficeIMO.Drawing.OfficeTransform.Identity,
                0D,
                effect => effect.Image(Png, 20D, 30D, 80D, 40D)))
            .ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(invisible);
        Assert.Single(Assert.Single(logical.Pages).Images);
        PdfHtmlConversionResult html = logical.ToHtmlResult(new PdfToHtmlOptions {
            Profile = PdfHtmlProfile.Semantic,
            IncludeImagePlaceholders = false
        });

        Assert.False(html.HasLoss);
        Assert.DoesNotContain(html.Report.Warnings, static warning =>
            warning.Code is "PdfImagesOmitted" or "PdfSemanticLayoutReflowed");

        PdfWordConversionResult word = logical.ToWordDocumentResult(new PdfToWordOptions {
            ImportImages = false,
            IncludeImagePlaceholders = false
        });
        using (word.Value) {
            Assert.False(word.HasLoss);
            Assert.Contains(word.Value.Paragraphs, static paragraph =>
                paragraph.Text == "No supported PDF content detected.");
            Assert.DoesNotContain(word.Report.Warnings, static warning =>
                warning.Code is "PdfImageSkipped" or "PdfEditableLayoutReconstructed");
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SuppressedImageOnlyPageDoesNotCreateAnEmptyWordPage(bool preserveSourcePageSize) {
        byte[] source = CreateDocument()
            .Paragraph(paragraph => paragraph.Text("Visible first page"))
            .PageBreak()
            .Canvas(canvas => canvas.Effect(
                OfficeIMO.Drawing.OfficeTransform.Identity,
                0D,
                effect => effect.Image(Png, 20D, 30D, 80D, 40D)))
            .ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);

        PdfWordConversionResult word = logical.ToWordDocumentResult(new PdfToWordOptions {
            PreserveSourcePageSize = preserveSourcePageSize,
            IncludeEmptyPages = false
        });
        using (word.Value) {
            Assert.Contains(word.Value.Paragraphs, static paragraph =>
                paragraph.Text.Contains("Visible first page", StringComparison.Ordinal));
            Assert.Empty(word.Value.Images);
            Assert.Empty(word.Value.PageBreaks);
            Assert.Single(word.Value.Sections);
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfInvisibleImagePlacementSuppressed" &&
                warning.LossKind == OfficeConversionLossKind.None);
        }
    }

    [Fact]
    public void UnplacedImageResourcesAreLossFreeInPositionedHtml() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(
            "q 40 0 0 20 20 30 cm /Im1 Do Q\n"));
        PdfLogicalPage page = Assert.Single(logical.Pages);
        PdfLogicalImage placedImage = Assert.Single(page.Images);
        var unplacedImage = new PdfLogicalImage(placedImage.SourceImage);
        var builder = new StringBuilder();
        PdfToHtmlOptions options = PdfToHtmlOptions.CreatePositionedReviewProfile();

        PdfHtmlConverterExtensions.AppendPositionedImagePlaceholders(
            builder,
            page,
            new[] { unplacedImage },
            options);

        Assert.DoesNotContain("pdf-image-placeholder", builder.ToString(), StringComparison.Ordinal);
        Assert.Equal(0, options.EmittedImagePlaceholderCount);
        Assert.False(options.Report.HasLoss);
        Assert.Contains(options.Report.Warnings, static warning =>
            warning.Code == "UnplacedImageResourceNotEmbedded" &&
            warning.LossKind == OfficeConversionLossKind.None);
    }

    [Theory]
    [InlineData("q 40 0 0 20 200 30 cm /Im1 Do Q\n")]
    [InlineData("q 0 0 10 10 re W n 40 0 0 20 80 80 cm /Im1 Do Q\n")]
    [InlineData("q 0 0 0 20 20 30 cm /Im1 Do Q\n")]
    public void ImagePlacementsWithoutVisiblePageIntersectionAreLossFree(string content) {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(content));
        Assert.Single(Assert.Single(logical.Pages).Images);
        Assert.Equal(0, PdfLogicalTableAnalysis.AnalyzeExtractionScope(logical).ImageCount);

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Empty(word.Value.Images);
            Assert.False(word.HasLoss);
            Assert.Contains(word.Value.Paragraphs, static paragraph =>
                paragraph.Text == "No supported PDF content detected.");
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfNonVisibleImagePlacementSuppressed" &&
                warning.LossKind == OfficeConversionLossKind.None);
        }

        PdfHtmlConversionResult html = logical.ToHtmlResult(PdfToHtmlOptions.CreateSemanticProfile());
        Assert.DoesNotContain("data:image/", html.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("<figure class=\"pdf-image-placeholder\"", html.Value, StringComparison.Ordinal);
        Assert.Equal(0, html.Summary.ImagePlaceholderCount);
        Assert.False(html.HasLoss);
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == "NonVisibleImagePlacementSuppressed" &&
            warning.LossKind == OfficeConversionLossKind.None);

        PdfPowerPointConversionResult editable = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using (editable.Value) {
            Assert.Empty(editable.Value.Slides.SelectMany(static slide => slide.Pictures));
            Assert.Equal(0, Assert.Single(editable.Report.EditablePages).OmittedImageCount);
            Assert.DoesNotContain(editable.Report.Warnings, static warning =>
                warning.LossKind == OfficeConversionLossKind.Omission);
            Assert.Contains(editable.Report.Warnings, static warning =>
                warning.Code == "PdfNonVisibleImagePlacementSuppressed" &&
                warning.LossKind == OfficeConversionLossKind.None);
        }

        PdfExcelTableImportResult excel = logical.ImportTablesToExcelDocumentResult();
        using (excel.Value) {
            Assert.False(excel.HasLoss);
            Assert.Equal(0, excel.Report.SourceScope.ImageCount);
        }

        PdfPowerPointConversionResult tables = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableTables());
        using (tables.Value) {
            Assert.False(tables.HasLoss);
            Assert.Equal(0, tables.Report.SourceScope.ImageCount);
            Assert.DoesNotContain(tables.Report.Warnings, static warning => warning.Code == "PdfImagesNotEditable");
        }
    }

    [Fact]
    public void NonVisibleUnsupportedImagePayloadDoesNotBecomePowerPointOmissionLoss() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(
            "q 40 0 0 20 200 30 cm /Im1 Do Q\n",
            imageMask: true));
        Assert.True(Assert.Single(Assert.Single(logical.Pages).Images).SourceImage.IsImageMask);

        PdfPowerPointConversionResult editable = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using (editable.Value) {
            Assert.Equal(0, Assert.Single(editable.Report.EditablePages).OmittedImageCount);
            Assert.DoesNotContain(editable.Report.Warnings, static warning =>
                warning.LossKind == OfficeConversionLossKind.Omission);
            Assert.Contains(editable.Report.Warnings, static warning =>
                warning.Code == "PdfNonVisibleImagePlacementSuppressed" &&
                warning.LossKind == OfficeConversionLossKind.None);
        }
    }

    [Fact]
    public void ClippedImagePlacementsNeverEmbedRawPixelsAcrossEditableAdapters() {
        byte[] source = CreateDocument()
            .Canvas(canvas => canvas.Clip(
                40D,
                40D,
                20D,
                20D,
                clipped => clipped.Image(Png, 20D, 20D, 80D, 60D)))
            .ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);
        Assert.NotNull(placement.Clip);

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Empty(word.Value.Images);
            Assert.True(word.HasLoss);
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfImageClipNotSafelyEditable" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }

        PdfHtmlConversionResult html = logical.ToHtmlResult(PdfToHtmlOptions.CreateSemanticProfile());
        Assert.DoesNotContain("data:image/", html.Value, StringComparison.Ordinal);
        Assert.True(html.HasLoss);
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == "ImageClipNotSafelyEditable" &&
            warning.LossKind == OfficeConversionLossKind.Omission);

        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using (powerPoint.Value) {
            Assert.Empty(powerPoint.Value.Slides.SelectMany(static slide => slide.Pictures));
            Assert.True(powerPoint.HasLoss);
            Assert.Contains(powerPoint.Report.Warnings, static warning =>
                warning.Code == "PdfImageClipNotSafelyEditable" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }
    }

    [Fact]
    public void TextClippedImagePlacementsNeverEmbedRawPixelsAcrossEditableAdapters() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateTextClippedRawImagePdf());
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);
        Assert.NotNull(placement.Clip);
        Assert.True(placement.Clip.ContainsTextClipping);
        Assert.True(placement.Clip.IsRectangle);
        Assert.True(placement.Clip.IsExact);

        AssertRawImageOmittedAcrossEditableAdapters(
            logical,
            "PdfImageClipNotSafelyEditable",
            "ImageClipNotSafelyEditable");
    }

    [Fact]
    public void PositionedPageAppearanceNeverEmbedsPixelsHiddenByImageClipping() {
        byte[] source = CreateDocument()
            .Canvas(canvas => canvas.Clip(
                40D,
                40D,
                20D,
                20D,
                clipped => clipped.Image(Png, 20D, 20D, 80D, 60D)))
            .ToBytes();

        PdfHtmlConversionResult html = PdfDocument.Load(source)
            .ToHtmlResult(PdfToHtmlOptions.CreatePositionedReviewProfile());

        Assert.DoesNotContain("pdf-page-appearance", html.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("data:image/", html.Value, StringComparison.Ordinal);
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == "PageAppearanceUnsafeImageFallback" &&
            warning.LossKind == OfficeConversionLossKind.None);
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == "ImageClipNotSafelyEditable" &&
            warning.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void PositionedPageAppearanceHashesRepeatedBackingImageOnlyOnce() {
        string content = string.Concat(Enumerable.Range(0, 16).Select(index =>
            "q 8 0 0 8 " + (8 + index % 4 * 20).ToString(System.Globalization.CultureInfo.InvariantCulture) + " " +
            (8 + index / 4 * 20).ToString(System.Globalization.CultureInfo.InvariantCulture) + " cm /Im1 Do Q\n"));
        byte[] source = CreateRawImagePdf(content);
        OfficeIMO.Drawing.OfficeDrawing drawing = PdfReadDocument.Open(source).Pages[0].ToDrawing();
        OfficeIMO.Drawing.OfficeDrawingImage[] drawingImages = drawing.Images.ToArray();
        Assert.Equal(16, drawingImages.Length);
        Assert.All(drawingImages, image => Assert.Same(drawingImages[0].EncodedBytes, image.EncodedBytes));

        int hashCount = 0;
        PdfHtmlConverterExtensions.ImagePayloadHashObserverForTesting = () => hashCount++;
        try {
            PdfHtmlConversionResult html = PdfDocument.Load(source)
                .ToHtmlResult(PdfToHtmlOptions.CreatePositionedReviewProfile());
            Assert.Contains("pdf-page-appearance", html.Value, StringComparison.Ordinal);
        } finally {
            PdfHtmlConverterExtensions.ImagePayloadHashObserverForTesting = null;
        }

        Assert.Equal(2, hashCount); // One logical-read payload and one separately parsed visual-source payload.
    }

    [Fact]
    public void PositionedPageAppearanceObservesCancellationBeforeHashingANewImagePayload() {
        byte[] source = CreateRawImagePdf("q 40 0 0 20 20 30 cm /Im1 Do Q\n");
        using var cancellation = new System.Threading.CancellationTokenSource();
        PdfHtmlConverterExtensions.ImagePayloadHashObserverForTesting = cancellation.Cancel;
        try {
            Assert.Throws<OperationCanceledException>(() => PdfDocument.Load(source)
                .ToHtmlResult(PdfToHtmlOptions.CreatePositionedReviewProfile(), cancellation.Token));
        } finally {
            PdfHtmlConverterExtensions.ImagePayloadHashObserverForTesting = null;
        }
    }

    [Fact]
    public void PositionedPageAppearanceChargesInterpolatedDecodedImagesToOneAggregatePixelBudget() {
        byte[] tiff = OfficeIMO.Drawing.OfficeRasterImageEncoder.Encode(
            new OfficeIMO.Drawing.OfficeRasterImage(300, 200, OfficeIMO.Drawing.OfficeColor.Red),
            OfficeIMO.Drawing.OfficeImageExportFormat.Tiff);
        PdfExtractedImage first = new PdfExtractedImage(
            1, "Im1", 5, 300, 200, 24, "DeviceRGB", string.Empty,
            tiff, "tiff", "image/tiff", isImageFile: true, interpolate: true);
        PdfExtractedImage second = new PdfExtractedImage(
            1, "Im2", 6, 300, 200, 24, "DeviceRGB", string.Empty,
            tiff, "tiff", "image/tiff", isImageFile: true, interpolate: true);

        Assert.True(PdfHtmlConverterExtensions.CanRenderPageAppearanceImagesWithinBudgetForTesting(
            new[] { first }, 100_000));
        Assert.False(PdfHtmlConverterExtensions.CanRenderPageAppearanceImagesWithinBudgetForTesting(
            new[] { first, second }, 100_000));
    }

    [Fact]
    public void RotatedPageClipsAreComparedInTheSameVisualCoordinateSpaceAsImages() {
        const string content = "q 0 0 150 60 re W n 40 0 0 80 20 20 cm /Im1 Do Q\n";
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(content, pageEntries: "/Rotate 90"));
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);
        Assert.NotNull(placement.Clip);

        AssertRawImageOmittedAcrossEditableAdapters(logical, "PdfImageClipNotSafelyEditable", "ImageClipNotSafelyEditable");
    }

    [Fact]
    public void RotatedPageClipThatContainsTheWholeImageRemainsEditable() {
        const string content = "q 0 0 80 120 re W n 40 0 0 80 20 20 cm /Im1 Do Q\n";
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(content, pageEntries: "/Rotate 90"));

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Single(word.Value.Images);
            Assert.DoesNotContain(word.Report.Warnings, static warning =>
                warning.Code == "PdfImageClipNotSafelyEditable");
        }
    }

    [Fact]
    public void UnsupportedImageBlendModesNeverEmbedRawPixelsAcrossEditableAdapters() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateImageWithGraphicsStatePdf("/BM /DefinitelyUnsupported"));
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);
        Assert.True(placement.HasUnsupportedBlendMode);
        Assert.False(placement.HasUnsupportedPaintState);

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Empty(word.Value.Images);
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfImageUnsupportedBlendModeNotSafelyEditable" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }

        PdfHtmlConversionResult html = logical.ToHtmlResult(PdfToHtmlOptions.CreateSemanticProfile());
        Assert.DoesNotContain("data:image/", html.Value, StringComparison.Ordinal);
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == "ImageUnsupportedBlendModeNotSafelyEditable" &&
            warning.LossKind == OfficeConversionLossKind.Omission);

        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using (powerPoint.Value) {
            Assert.Empty(powerPoint.Value.Slides.SelectMany(static slide => slide.Pictures));
            Assert.Contains(powerPoint.Report.Warnings, static warning =>
                warning.Code == "PdfImageUnsupportedBlendModeNotSafelyEditable" &&
                warning.LossKind == OfficeConversionLossKind.Omission);
        }
    }

    [Fact]
    public void NonBlendGraphicsStateEntriesDoNotMasqueradeAsUnsupportedBlendModes() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateImageWithGraphicsStatePdf("/SA true /SM 0.02 /AIS false"));
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);

        Assert.False(placement.HasUnsupportedBlendMode);
        Assert.True(placement.HasUnsupportedPaintState);
        Assert.False(placement.HasUnsupportedImagePaintEffect);
        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Single(word.Value.Images);
            Assert.DoesNotContain(word.Report.Warnings, static warning =>
                warning.Code == "PdfImageUnsupportedBlendModeNotSafelyEditable");
        }
    }

    [Fact]
    public void NullBlendModeDoesNotCreateFalseImageLoss() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateImageWithGraphicsStatePdf("/BM null"));
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);

        Assert.False(placement.HasUnsupportedBlendMode);
        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Single(word.Value.Images);
            Assert.DoesNotContain(word.Report.Warnings, static warning =>
                warning.Code == "PdfImageUnsupportedBlendModeNotSafelyEditable");
        }
    }

    [Fact]
    public void NullBlendModePreservesAnInheritedNonNormalBlendMode() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(
            "q /GS1 gs /GS2 gs 80 0 0 40 20 30 cm /Im1 Do Q\n",
            "/BM /Multiply",
            secondGraphicsStateEntries: "/BM null"));
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);

        Assert.Equal(OfficeIMO.Drawing.OfficeBlendMode.Multiply, placement.BlendMode);
        Assert.False(placement.HasUnsupportedBlendMode);
    }

    [Theory]
    [InlineData("/TR << /FunctionType 2 /Domain [0 1] /C0 [0] /C1 [0] /N 1 >>")]
    [InlineData("/op true /OPM 1")]
    [InlineData("/AIS true /ca 0.5")]
    public void ImagePaintEffectsNeverExposeUntransformedRawPixelsAcrossEditableAdapters(string graphicsStateEntries) {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateImageWithGraphicsStatePdf(graphicsStateEntries));
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);

        Assert.True(placement.HasUnsupportedPaintState);
        Assert.True(placement.HasUnsupportedImagePaintEffect);
        AssertRawImageOmittedAcrossEditableAdapters(logical, "PdfImagePaintEffectNotSafelyEditable", "ImagePaintEffectNotSafelyEditable");
    }

    [Fact]
    public void ExplicitImagePaintEffectResetReplacesThePriorGraphicsState() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(
            "q /GS1 gs /GS2 gs 80 0 0 40 20 30 cm /Im1 Do Q\n",
            "/op true /OPM 1 /AIS true",
            secondGraphicsStateEntries: "/op false /OPM 0 /AIS false"));
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);

        Assert.False(placement.HasUnsupportedImagePaintEffect);
        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Single(word.Value.Images);
            Assert.DoesNotContain(word.Report.Warnings, static warning =>
                warning.Code == "PdfImagePaintEffectNotSafelyEditable");
        }
    }

    [Theory]
    [InlineData("/TR << /FunctionType 2 /Domain [0 1] /C0 [0] /C1 [0] /N 1 >>", "/TR2 /Identity")]
    [InlineData("/BG << /FunctionType 2 /Domain [0 1] /C0 [0] /C1 [0] /N 1 >>", "/BG2 /Default")]
    [InlineData("/UCR << /FunctionType 2 /Domain [0 1] /C0 [0] /C1 [0] /N 1 >>", "/UCR2 /Default")]
    [InlineData("/HT << /HalftoneType 5 >>", "/HT /Default")]
    public void NamedImagePaintEffectResetReplacesThePriorGraphicsState(
        string initialGraphicsState,
        string resetGraphicsState) {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(
            "q /GS1 gs /GS2 gs 80 0 0 40 20 30 cm /Im1 Do Q\n",
            initialGraphicsState,
            secondGraphicsStateEntries: resetGraphicsState));
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);

        Assert.False(placement.HasUnsupportedImagePaintEffect);
    }

    [Fact]
    public void PartialImagePaintEffectResetPreservesOtherActiveEffects() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(
            "q /GS1 gs /GS2 gs 80 0 0 40 20 30 cm /Im1 Do Q\n",
            "/op true /AIS true /ca 0.5",
            secondGraphicsStateEntries: "/op false"));
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);

        Assert.True(placement.HasUnsupportedImagePaintEffect);
        AssertRawImageOmittedAcrossEditableAdapters(logical, "PdfImagePaintEffectNotSafelyEditable", "ImagePaintEffectNotSafelyEditable");
    }

    [Fact]
    public void OverprintModeWithoutActiveOverprintDoesNotCreateFalseImageLoss() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateImageWithGraphicsStatePdf("/OPM 1"));
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);

        Assert.False(placement.HasUnsupportedImagePaintEffect);
        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Single(word.Value.Images);
        }
    }

    [Fact]
    public void AlphaIsShapeWithoutActiveTransparencyDoesNotCreateFalseImageLoss() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateImageWithGraphicsStatePdf("/AIS true"));
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);

        Assert.False(placement.HasUnsupportedImagePaintEffect);
        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Single(word.Value.Images);
            Assert.DoesNotContain(word.Report.Warnings, static warning =>
                warning.Code == "PdfImagePaintEffectNotSafelyEditable");
        }
    }

    [Theory]
    [InlineData("/op true", "/op null")]
    [InlineData("/TR << /FunctionType 2 /Domain [0 1] /C0 [0] /C1 [0] /N 1 >>", "/TR2 null")]
    [InlineData("/AIS true /ca 0.5", "/AIS null")]
    public void NullGraphicsStateEntriesPreserveInheritedImagePaintEffects(
        string initialGraphicsState,
        string nullGraphicsState) {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(
            "q /GS1 gs /GS2 gs 80 0 0 40 20 30 cm /Im1 Do Q\n",
            initialGraphicsState,
            secondGraphicsStateEntries: nullGraphicsState));
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);

        Assert.True(placement.HasUnsupportedImagePaintEffect);
        AssertRawImageOmittedAcrossEditableAdapters(
            logical,
            "PdfImagePaintEffectNotSafelyEditable",
            "ImagePaintEffectNotSafelyEditable");
    }

    [Fact]
    public void NestedFormCanExplicitlyResetInheritedImagePaintEffect() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateNestedFormImageWithGraphicsStatePdf(
            outerGraphicsStateEntries: "/op true",
            innerGraphicsStateEntries: "/op false"));
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);

        Assert.False(placement.HasUnsupportedImagePaintEffect);
        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Single(word.Value.Images);
            Assert.DoesNotContain(word.Report.Warnings, static warning =>
                warning.Code == "PdfImagePaintEffectNotSafelyEditable");
        }
    }

    [Fact]
    public void RetainedOverprintModeBecomesLossBearingWhenOverprintIsEnabledLater() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(
            "q /GS1 gs /GS2 gs 80 0 0 40 20 30 cm /Im1 Do Q\n",
            "/OPM 1",
            secondGraphicsStateEntries: "/op true"));
        PdfImagePlacement placement = Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements);

        Assert.True(placement.HasUnsupportedImagePaintEffect);
        AssertRawImageOmittedAcrossEditableAdapters(logical, "PdfImagePaintEffectNotSafelyEditable", "ImagePaintEffectNotSafelyEditable");
    }

    [Fact]
    public void FloatingPositionAndNaturalImageSizeCanBeSelectedIndependently() {
        byte[] source = CreateDocument()
            .Canvas(canvas => canvas.Image(Png, 20D, 30D, 80D, 40D))
            .ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);

        PdfWordConversionResult word = logical.ToWordDocumentResult(new PdfToWordOptions {
            PreserveImagePlacementPosition = true,
            PreserveImagePlacementSize = false
        });
        using (word.Value) {
            OfficeIMO.Word.WordImage image = Assert.Single(word.Value.Images);
            Assert.Equal(OfficeIMO.Word.WordImageTextWrapping.InFrontOfText, image.WrapText);
            Assert.Equal(8D, image.Width);
            Assert.Equal(4D, image.Height);
            OfficeIMO.Word.WordParagraph anchor = Assert.Single(word.Value.Paragraphs);
            Assert.Equal(OfficeIMO.Word.WordLineSpacingRule.Exact, anchor.LineSpacingRule);
            Assert.Equal(1, anchor.LineSpacing);
        }
    }

    [Fact]
    public void PreservedEditableGeometryUsesThePageUserUnitScale() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(
            "q 80 0 0 40 20 30 cm /Im1 Do Q\n",
            pageEntries: "/UserUnit 2"));

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            OfficeIMO.Word.WordImage image = Assert.Single(word.Value.Images);
            Assert.Equal(80D * 2D * 96D / 72D, image.Width!.Value, 6);
            Assert.Equal(40D * 2D * 96D / 72D, image.Height!.Value, 6);
        }

        PdfHtmlConversionResult html = logical.ToHtmlResult(PdfToHtmlOptions.CreatePositionedReviewProfile());
        Assert.Contains("style=\"width:320pt;height:320pt;\"", html.Value, StringComparison.Ordinal);
        Assert.Contains("width:160pt;height:80pt;", html.Value, StringComparison.Ordinal);
    }

    [Fact]
    public void WordFallsBackToInlineImagesWhenSourcePageSizeCannotBeApplied() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(
            "q 80 0 0 40 1700 30 cm /Im1 Do Q\n",
            mediaBox: "[0 0 2000 800]"));

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            using var serialized = new MemoryStream(word.Value.ToBytes());
            using OfficeIMO.Word.WordDocument reopened = OfficeIMO.Word.WordDocument.Load(serialized);
            OfficeIMO.Word.WordImage image = Assert.Single(reopened.Images);
            Assert.Equal(OfficeIMO.Word.WordImageTextWrapping.InLineWithText, image.WrapText);
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfSourcePageSizeNotApplied");
        }
    }

    [Fact]
    public void WordDoesNotUsePageRelativeImageAnchorsWhenSourcePageSizingIsDisabled() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(
            "q 80 0 0 40 20 30 cm /Im1 Do Q\n"));

        PdfWordConversionResult word = logical.ToWordDocumentResult(new PdfToWordOptions {
            PreserveSourcePageSize = false
        });
        using (word.Value) {
            using var serialized = new MemoryStream(word.Value.ToBytes());
            using OfficeIMO.Word.WordDocument reopened = OfficeIMO.Word.WordDocument.Load(serialized);
            Assert.Equal(
                OfficeIMO.Word.WordImageTextWrapping.InLineWithText,
                Assert.Single(reopened.Images).WrapText);
        }
    }

    [Fact]
    public void PreservedWordPageSizeScalesEditableTypographyByThePageUserUnit() {
        const string marker = "Scaled user unit text";
        byte[] source = WithUserUnit(
            CreateDocument()
                .Paragraph(paragraph => paragraph.FontSize(10D).Text(marker))
                .ToBytes(),
            2D);
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            OfficeIMO.Word.WordParagraph paragraph = Assert.Single(
                word.Value.Paragraphs,
                candidate => candidate.Text.Contains(marker, StringComparison.Ordinal));
            Assert.Equal(20D, paragraph.FontSizePoints);
        }

        PdfWordConversionResult unscaled = logical.ToWordDocumentResult(new PdfToWordOptions {
            PreserveSourcePageSize = false
        });
        using (unscaled.Value) {
            OfficeIMO.Word.WordParagraph paragraph = Assert.Single(
                unscaled.Value.Paragraphs,
                candidate => candidate.Text.Contains(marker, StringComparison.Ordinal));
            Assert.Equal(10D, paragraph.FontSizePoints);
        }
    }

    [Fact]
    public void WordDoesNotScaleLaterPageTypographyWhenPageBreaksAreNotPreserved() {
        const string firstMarker = "First page typography";
        const string secondMarker = "Second page typography";
        byte[] source = WithPageUserUnit(
            CreateDocument()
                .Paragraph(paragraph => paragraph.FontSize(10D).Text(firstMarker))
                .PageBreak()
                .Paragraph(paragraph => paragraph.FontSize(10D).Text(secondMarker))
                .ToBytes(),
            pageNumber: 2,
            userUnit: 2D);
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);

        PdfWordConversionResult word = logical.ToWordDocumentResult(new PdfToWordOptions {
            PreservePageBreaks = false,
            PreserveSourcePageSize = true
        });
        using (word.Value) {
            Assert.Equal(10D, Assert.Single(word.Value.Paragraphs,
                paragraph => paragraph.Text.Contains(firstMarker, StringComparison.Ordinal)).FontSizePoints);
            Assert.Equal(10D, Assert.Single(word.Value.Paragraphs,
                paragraph => paragraph.Text.Contains(secondMarker, StringComparison.Ordinal)).FontSizePoints);
            Assert.Single(word.Value.Sections);
        }
    }

    [Fact]
    public void PreservedWordPageSizeScalesImportedTableCellTypographyByThePageUserUnit() {
        byte[] source = WithUserUnit(
            CreateDocument()
                .Table(new[] {
                    new[] { "Item", "Amount" },
                    new[] { "Service", "42.00" }
                })
                .ToBytes(),
            2D);
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);
        Assert.NotEmpty(logical.Tables);

        PdfWordConversionResult word = logical.ToWordDocumentResult(PdfToWordOptions.CreateTablesOnly());
        using (word.Value) {
            using var serialized = new MemoryStream(word.Value.ToBytes());
            using OfficeIMO.Word.WordDocument reopened = OfficeIMO.Word.WordDocument.Load(serialized);
            OfficeIMO.Word.WordTable table = Assert.Single(reopened.Tables);
            Assert.All(table.Rows.SelectMany(static row => row.Cells), static cell =>
                Assert.All(cell.Paragraphs, static paragraph => Assert.Equal(22D, paragraph.FontSizePoints)));
        }
    }

    [Theory]
    [InlineData("q -80 0 0 40 100 30 cm /Im1 Do Q\n", true, false)]
    [InlineData("q 80 0 0 -40 20 70 cm /Im1 Do Q\n", false, true)]
    [InlineData("q -80 0 0 -40 100 70 cm /Im1 Do Q\n", true, true)]
    public void WordPreservesAxisAlignedImageReflections(
        string content,
        bool expectedHorizontalFlip,
        bool expectedVerticalFlip) {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(content));

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            using var serialized = new MemoryStream(word.Value.ToBytes());
            using OfficeIMO.Word.WordDocument reopened = OfficeIMO.Word.WordDocument.Load(serialized);
            OfficeIMO.Word.WordImage image = Assert.Single(reopened.Images);
            Assert.Equal(expectedHorizontalFlip, image.HorizontalFlip);
            Assert.Equal(expectedVerticalFlip, image.VerticalFlip);
        }
    }

    [Fact]
    public void WordFloatingImagesUsePdfPaintOrderForTheirZOrder() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateTwoRawImagePdf(
            "q 60 0 0 30 30 40 cm /Im2 Do Q\n" +
            "q 80 0 0 40 20 30 cm /Im1 Do Q\n"));
        PdfLogicalPage page = Assert.Single(logical.Pages);
        Assert.Equal(2, page.Images.Count);

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            using var serialized = new MemoryStream(word.Value.ToBytes());
            using OfficeIMO.Word.WordDocument reopened = OfficeIMO.Word.WordDocument.Load(serialized);
            Assert.Equal(2, reopened.Images.Count);
            OfficeIMO.Word.WordImage front = Assert.Single(reopened.Images, static image => image.Width > 100D);
            OfficeIMO.Word.WordImage back = Assert.Single(reopened.Images, static image => image.Width < 100D);
            Assert.True(front.ZOrder > back.ZOrder);
        }
    }

    [Fact]
    public void WordRotatesInlineImageWhenPageRotationSwapsVisualDimensions() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(
            "q 80 0 0 40 20 30 cm /Im1 Do Q\n",
            pageEntries: "/Rotate 90"));

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            OfficeIMO.Word.WordImage image = Assert.Single(word.Value.Images);
            Assert.Equal(90, image.Rotation);
            Assert.Equal(80D * 96D / 72D, image.Width!.Value, 6);
            Assert.Equal(40D * 96D / 72D, image.Height!.Value, 6);
        }
    }

    [Theory]
    [InlineData(90)]
    [InlineData(180)]
    [InlineData(270)]
    public void PowerPointRotatesEditableImageWithThePdfPage(int pageRotation) {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(
            "q 80 0 0 40 20 30 cm /Im1 Do Q\n",
            pageEntries: "/Rotate " + pageRotation.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        PdfLogicalPage page = Assert.Single(logical.Pages);
        PdfVisualBounds visual = page.TransformBoundsToVisual(20D, 30D, 100D, 70D);

        PdfPowerPointConversionResult result = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        byte[] bytes;
        using (result.Value) {
            bytes = result.Value.ToBytes();
        }

        using var presentationStream = new MemoryStream(bytes);
        using OfficeIMO.PowerPoint.PowerPointPresentation presentation =
            OfficeIMO.PowerPoint.PowerPointPresentation.Load(presentationStream);
        OfficeIMO.PowerPoint.PowerPointPicture picture = Assert.Single(Assert.Single(presentation.Slides).Pictures);
        double scale = presentation.SlideSize.WidthPoints / 160D;
        Assert.Equal(pageRotation, picture.Rotation);
        Assert.Equal(80D * scale, picture.WidthPoints, 4);
        Assert.Equal(40D * scale, picture.HeightPoints, 4);
        Assert.Equal((visual.Left + visual.Width / 2D) * scale, picture.LeftPoints + picture.WidthPoints / 2D, 4);
        Assert.Equal((visual.Top + visual.Height / 2D) * scale, picture.TopPoints + picture.HeightPoints / 2D, 4);
    }

    [Fact]
    public void FloatingImagePaintedBeforeOverlappingTextStaysBehindTheText() {
        byte[] source = CreateDocument()
            .Canvas(canvas => canvas
                .Image(Png, 20D, 30D, 120D, 80D)
                .Text("Foreground invoice text", 30D, 50D, 100D, 24D, fontSize: 14D))
            .ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);
        PdfLogicalPage page = Assert.Single(logical.Pages);
        PdfImagePlacement placement = Assert.Single(Assert.Single(page.Images).Placements);
        PdfTextSpan[] textSpans = page.TextBlocks.SelectMany(static block => block.Spans).ToArray();
        Assert.NotEmpty(textSpans);
        Assert.All(textSpans, text => Assert.True(placement.PaintOrder < text.PaintOrder));

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Equal(
                OfficeIMO.Word.WordImageTextWrapping.BehindText,
                Assert.Single(word.Value.Images).WrapText);
        }
    }

    [Fact]
    public void FloatingImagePaintedAfterOverlappingTextStaysInFrontOfTheText() {
        byte[] source = CreateDocument()
            .Canvas(canvas => canvas
                .Text("Background invoice text", 30D, 50D, 100D, 24D, fontSize: 14D)
                .Image(Png, 20D, 30D, 120D, 80D))
            .ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);
        PdfLogicalPage page = Assert.Single(logical.Pages);
        PdfImagePlacement placement = Assert.Single(Assert.Single(page.Images).Placements);
        PdfTextSpan[] textSpans = page.TextBlocks.SelectMany(static block => block.Spans).ToArray();
        Assert.NotEmpty(textSpans);
        Assert.All(textSpans, text => Assert.True(text.PaintOrder < placement.PaintOrder));

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Equal(
                OfficeIMO.Word.WordImageTextWrapping.InFrontOfText,
                Assert.Single(word.Value.Images).WrapText);
        }
    }

    [Fact]
    public void NonDefaultImageOpacityIsMappedAcrossEditableAdapters() {
        byte[] source = CreateDocument()
            .Canvas(canvas => canvas.Effect(
                OfficeIMO.Drawing.OfficeTransform.Identity,
                0.5D,
                effect => effect.Image(Png, 20D, 30D, 80D, 40D)))
            .ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Equal(50, Assert.Single(word.Value.Images).Transparency);
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfImageOpacityMapped" &&
                warning.LossKind == OfficeConversionLossKind.None);
        }

        PdfHtmlConversionResult html = logical.ToHtmlResult(PdfToHtmlOptions.CreateSemanticProfile());
        Assert.Contains("style=\"opacity:0.5;\"", html.Value, StringComparison.Ordinal);
        Assert.Contains(html.Report.Warnings, static warning =>
            warning.Code == "ImageOpacityMapped" &&
            warning.LossKind == OfficeConversionLossKind.None);

        PdfToHtmlOptions semanticPlaceholderOptions = PdfToHtmlOptions.CreateSemanticProfile();
        semanticPlaceholderOptions.ImageExportMode = PdfHtmlImageExportMode.PlaceholderOnly;
        PdfHtmlConversionResult semanticPlaceholder = logical.ToHtmlResult(semanticPlaceholderOptions);
        Assert.DoesNotContain(semanticPlaceholder.Report.Warnings, static warning =>
            warning.Code is "ImageOpacityMapped" or "ImageBlendModeApproximated");

        PdfToHtmlOptions positionedPlaceholderOptions = PdfToHtmlOptions.CreatePositionedReviewProfile();
        positionedPlaceholderOptions.ImageExportMode = PdfHtmlImageExportMode.PlaceholderOnly;
        PdfHtmlConversionResult positionedPlaceholder = logical.ToHtmlResult(positionedPlaceholderOptions);
        Assert.DoesNotContain(positionedPlaceholder.Report.Warnings, static warning =>
            warning.Code is "ImageOpacityMapped" or "ImageBlendModeApproximated");

        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using (powerPoint.Value) {
            Assert.Equal(50, Assert.Single(Assert.Single(powerPoint.Value.Slides).Pictures).FillTransparency);
            Assert.Contains(powerPoint.Report.Warnings, static warning =>
                warning.Code == "PdfImageOpacityMapped" &&
                warning.LossKind == OfficeConversionLossKind.None);
        }
    }

    [Fact]
    public void VeryLowNonzeroImageOpacityIsReportedWhenOfficePrecisionOmitsIt() {
        byte[] source = CreateRawImagePdf(
            "q /GS1 gs 80 0 0 40 20 30 cm /Im1 Do Q\n",
            "/ca 0.0000004");
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);
        Assert.Equal(0.0000004D, Assert.Single(Assert.Single(logical.Pages).Images).Placements[0].Opacity, 10);

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Equal(100, Assert.Single(word.Value.Images).Transparency);
            Assert.Contains(word.Report.Warnings, static warning =>
                warning.Code == "PdfImageOpacityMapped" &&
                warning.LossKind == OfficeConversionLossKind.Omission &&
                warning.Details["MappedTransparencyPercent"] == "100");
            Assert.True(word.Report.HasLoss);
            Assert.Throws<InvalidOperationException>(() => word.RequireNoLoss());
        }

        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using (powerPoint.Value) {
            Assert.Equal(100, Assert.Single(Assert.Single(powerPoint.Value.Slides).Pictures).FillTransparency);
            Assert.Contains(powerPoint.Report.Warnings, static warning =>
                warning.Code == "PdfImageOpacityMapped" &&
                warning.LossKind == OfficeConversionLossKind.Omission &&
                warning.Details["MappedTransparencyPercent"] == "100");
            Assert.True(powerPoint.Report.HasLoss);
            Assert.Throws<InvalidOperationException>(() => powerPoint.RequireNoLoss());
        }

        PdfHtmlConversionResult semanticHtml = logical.ToHtmlResult(PdfToHtmlOptions.CreateSemanticProfile());
        Assert.Contains("opacity:4E-07;", semanticHtml.Value, StringComparison.Ordinal);
        Assert.Contains(semanticHtml.Report.Warnings, static warning =>
            warning.Code == "ImageOpacityMapped" &&
            warning.LossKind == OfficeConversionLossKind.None);

        PdfHtmlConversionResult positionedHtml = logical.ToHtmlResult(PdfToHtmlOptions.CreatePositionedReviewProfile());
        Assert.Contains("opacity:4E-07;", positionedHtml.Value, StringComparison.Ordinal);
        Assert.Contains(positionedHtml.Report.Warnings, static warning =>
            warning.Code == "ImageOpacityMapped" &&
            warning.LossKind == OfficeConversionLossKind.None);
    }

    [Fact]
    public void PowerPointDoesNotReportMappedEffectsWhenImagePayloadIsOmitted() {
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(CreateRawImagePdf(
            "q /GS1 gs 80 0 0 40 20 30 cm /Im1 Do Q\n",
            "/ca 0.5",
            imageMask: true));
        Assert.Equal(0.5D, Assert.Single(Assert.Single(Assert.Single(logical.Pages).Images).Placements).Opacity);

        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using (powerPoint.Value) {
            Assert.Empty(powerPoint.Value.Slides.SelectMany(static slide => slide.Pictures));
            Assert.DoesNotContain(powerPoint.Report.Warnings, static warning =>
                warning.Code is "PdfImageOpacityMapped" or "PdfImageBlendModeApproximated");
        }
    }

    [Fact]
    public void Jpeg2000WithUnappliedDecodeNeverExposesRawPixelsAcrossEditableAdapters() {
        byte[] jpx = File.ReadAllBytes(Path.Combine(
            AppContext.BaseDirectory,
            "Pdf",
            "Fixtures",
            "Interoperability",
            "Scans",
            "red-rgb.jp2"));
        byte[] source = CreateRawImagePdf(
            "q 80 0 0 40 20 30 cm /Im1 Do Q\n",
            imageBytes: jpx,
            imageDefinition: "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /JPXDecode /Decode [1 0 1 0 1 0]");
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);
        PdfLogicalImage image = Assert.Single(Assert.Single(logical.Pages).Images);
        Assert.True(image.SourceImage.IsImageFile);
        Assert.True(image.SourceImage.HasExplicitDecode);
        Assert.True(image.SourceImage.HasUnsafePassThroughDecode);
        Assert.Equal("image/jp2", image.SourceImage.MimeType);

        AssertRawImageOmittedAcrossEditableAdapters(
            logical,
            "PdfImageDecodeNotSafelyEditable",
            "ImageDecodeNotSafelyEditable");

        PdfHtmlConversionResult positioned = PdfDocument.Load(source)
            .ToHtmlResult(PdfToHtmlOptions.CreatePositionedReviewProfile());
        Assert.DoesNotContain("data:image/", positioned.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("pdf-page-appearance", positioned.Value, StringComparison.Ordinal);
        Assert.Contains(positioned.Report.Warnings, static warning =>
            warning.Code == "PageAppearanceUnsafeImageFallback" &&
            warning.LossKind == OfficeConversionLossKind.None);
    }

    [Fact]
    public void Jpeg2000IdentityDecodeRemainsPortableAcrossEditableAdapters() {
        byte[] jpx = File.ReadAllBytes(Path.Combine(
            AppContext.BaseDirectory,
            "Pdf",
            "Fixtures",
            "Interoperability",
            "Scans",
            "red-rgb.jp2"));
        byte[] source = CreateRawImagePdf(
            "q 80 0 0 40 20 30 cm /Im1 Do Q\n",
            imageBytes: jpx,
            imageDefinition: "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /JPXDecode /Decode [0 1 0 1 0 1]");
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);
        PdfLogicalImage image = Assert.Single(Assert.Single(logical.Pages).Images);
        Assert.True(image.SourceImage.IsImageFile);
        Assert.True(image.SourceImage.HasExplicitDecode);
        Assert.False(image.SourceImage.HasUnsafePassThroughDecode);
        Assert.Equal(OfficeIMO.Drawing.OfficeImageFormat.Jpeg2000, OfficeIMO.Drawing.OfficeImageReader.Identify(jpx).Format);

        PdfHtmlConversionResult html = logical.ToHtmlResult(PdfToHtmlOptions.CreateSemanticProfile());
        Assert.Contains("data:image/jp2;base64,", html.Value, StringComparison.Ordinal);
        Assert.DoesNotContain(html.Report.Warnings, static warning => warning.Code == "ImageDecodeNotSafelyEditable");

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            using var serialized = new MemoryStream(word.Value.ToBytes());
            using OfficeIMO.Word.WordDocument reopened = OfficeIMO.Word.WordDocument.Load(serialized);
            Assert.Single(reopened.Images);
            Assert.Equal(jpx, Assert.Single(reopened.GetImageBytes()));
            Assert.Empty(reopened.ValidateDocument());
            Assert.DoesNotContain(word.Report.Warnings, static warning => warning.Code == "PdfImageDecodeNotSafelyEditable");
        }

        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using (powerPoint.Value) {
            using var serialized = new MemoryStream(powerPoint.Value.ToBytes());
            using OfficeIMO.PowerPoint.PowerPointPresentation reopened = OfficeIMO.PowerPoint.PowerPointPresentation.Load(serialized);
            OfficeIMO.PowerPoint.PowerPointPicture picture = Assert.Single(
                reopened.Slides.SelectMany(static slide => slide.Pictures));
            Assert.Equal("image/jp2", picture.ContentType);
            Assert.Equal(jpx, picture.GetImageBytes());
            Assert.Empty(reopened.ValidateDocument());
            Assert.DoesNotContain(powerPoint.Report.Warnings, static warning => warning.Code == "PdfImageDecodeNotSafelyEditable");
        }
    }

    [Fact]
    public void RawJpeg2000CodestreamIsNotEmbeddedAsAJP2FileAcrossEditableAdapters() {
        byte[] container = File.ReadAllBytes(Path.Combine(
            AppContext.BaseDirectory,
            "Pdf",
            "Fixtures",
            "Interoperability",
            "Scans",
            "red-rgb.jp2"));
        int codestream = Enumerable.Range(0, container.Length - 3).Single(index =>
            container[index] == 0xFF && container[index + 1] == 0x4F &&
            container[index + 2] == 0xFF && container[index + 3] == 0x51);
        byte[] rawCodestream = container.Skip(codestream).ToArray();
        byte[] source = CreateRawImagePdf(
            "q 80 0 0 40 20 30 cm /Im1 Do Q\n",
            imageBytes: rawCodestream,
            imageDefinition: "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /JPXDecode");
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);
        PdfLogicalImage image = Assert.Single(Assert.Single(logical.Pages).Images);
        Assert.False(image.SourceImage.IsImageFile);

        AssertJpeg2000PayloadNotEmbeddedAcrossEditableAdapters(logical);
    }

    [Fact]
    public void MalformedJp2ContainerIsNotEmbeddedAcrossEditableAdapters() {
        byte[] container = File.ReadAllBytes(Path.Combine(
            AppContext.BaseDirectory,
            "Pdf",
            "Fixtures",
            "Interoperability",
            "Scans",
            "red-rgb.jp2"));
        int tilePart = Enumerable.Range(0, container.Length - 1).First(index =>
            container[index] == 0xFF && container[index + 1] == 0x90);
        container[tilePart + 11] = 2; // TNsot declares an absent second tile-part.
        byte[] source = CreateRawImagePdf(
            "q 80 0 0 40 20 30 cm /Im1 Do Q\n",
            imageBytes: container,
            imageDefinition: "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /JPXDecode");
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);
        PdfLogicalImage image = Assert.Single(Assert.Single(logical.Pages).Images);
        Assert.False(image.SourceImage.IsImageFile);

        AssertJpeg2000PayloadNotEmbeddedAcrossEditableAdapters(logical);
    }

    [Fact]
    public void Jpeg2000NullDecodeDoesNotTriggerUnappliedDecodeLoss() {
        byte[] jpx = File.ReadAllBytes(Path.Combine(
            AppContext.BaseDirectory,
            "Pdf",
            "Fixtures",
            "Interoperability",
            "Scans",
            "red-rgb.jp2"));
        byte[] source = CreateRawImagePdf(
            "q 80 0 0 40 20 30 cm /Im1 Do Q\n",
            imageBytes: jpx,
            imageDefinition: "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /JPXDecode /Decode null");
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);
        PdfLogicalImage image = Assert.Single(Assert.Single(logical.Pages).Images);
        Assert.False(image.SourceImage.HasExplicitDecode);
        Assert.False(image.SourceImage.HasUnsafePassThroughDecode);

        PdfHtmlConversionResult html = logical.ToHtmlResult(PdfToHtmlOptions.CreateSemanticProfile());
        Assert.Contains("data:image/jp2;base64,", html.Value, StringComparison.Ordinal);
        Assert.DoesNotContain(html.Report.Warnings, static warning => warning.Code == "ImageDecodeNotSafelyEditable");

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.DoesNotContain(word.Report.Warnings, static warning => warning.Code == "PdfImageDecodeNotSafelyEditable");
        }

        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using (powerPoint.Value) {
            Assert.DoesNotContain(powerPoint.Report.Warnings, static warning => warning.Code == "PdfImageDecodeNotSafelyEditable");
        }
    }

    private static PdfDocument CreateDocument() => PdfDocument.Create(new PdfOptions {
        PageWidth = 160D,
        PageHeight = 160D,
        MarginLeft = 0D,
        MarginRight = 0D,
        MarginTop = 0D,
        MarginBottom = 0D
    });

    private static void AssertRawImageOmittedAcrossEditableAdapters(
        PdfDocumentReadResult logical,
        string officeWarningCode,
        string htmlWarningCode) {
        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Empty(word.Value.Images);
            Assert.Contains(word.Report.Warnings, warning =>
                warning.Code == officeWarningCode && warning.LossKind == OfficeConversionLossKind.Omission);
        }

        PdfHtmlConversionResult html = logical.ToHtmlResult(PdfToHtmlOptions.CreateSemanticProfile());
        Assert.DoesNotContain("data:image/", html.Value, StringComparison.Ordinal);
        Assert.Contains(html.Report.Warnings, warning =>
            warning.Code == htmlWarningCode && warning.LossKind == OfficeConversionLossKind.Omission);

        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using (powerPoint.Value) {
            Assert.Empty(powerPoint.Value.Slides.SelectMany(static slide => slide.Pictures));
            Assert.Contains(powerPoint.Report.Warnings, warning =>
                warning.Code == officeWarningCode && warning.LossKind == OfficeConversionLossKind.Omission);
        }
    }

    private static void AssertJpeg2000PayloadNotEmbeddedAcrossEditableAdapters(PdfDocumentReadResult logical) {
        PdfHtmlConversionResult html = logical.ToHtmlResult(PdfToHtmlOptions.CreateSemanticProfile());
        Assert.DoesNotContain("data:image/jp2", html.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("data:image/j2c", html.Value, StringComparison.Ordinal);
        Assert.True(html.HasLoss);

        PdfWordConversionResult word = logical.ToWordDocumentResult();
        using (word.Value) {
            Assert.Empty(word.Value.Images);
            Assert.True(word.HasLoss);
        }

        PdfPowerPointConversionResult powerPoint = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using (powerPoint.Value) {
            Assert.Empty(powerPoint.Value.Slides.SelectMany(static slide => slide.Pictures));
            Assert.True(powerPoint.HasLoss);
        }
    }

    private static byte[] CreateImageWithGraphicsStatePdf(string graphicsStateEntries) =>
        CreateRawImagePdf(
            "q /GS1 gs 80 0 0 40 20 30 cm /Im1 Do Q\n",
            graphicsStateEntries);

    private static byte[] WithUserUnit(byte[] source, double userUnit) =>
        PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            PdfIndirectObject page = Assert.Single(objects.Values, static item =>
                item.Value is PdfDictionary dictionary &&
                string.Equals(dictionary.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal));
            Assert.IsType<PdfDictionary>(page.Value).Items["UserUnit"] = new PdfNumber(userUnit);
            return security.InfoObjectNumber;
        });

    private static byte[] WithPageUserUnit(byte[] source, int pageNumber, double userUnit) =>
        PdfDocumentObjectGraphRewriter.Rewrite(source, null, null, (objects, security) => {
            PdfIndirectObject[] pages = objects.Values
                .Where(static item => item.Value is PdfDictionary dictionary &&
                    string.Equals(dictionary.Get<PdfName>("Type")?.Name, "Page", StringComparison.Ordinal))
                .OrderBy(static item => item.ObjectNumber)
                .ToArray();
            Assert.InRange(pageNumber, 1, pages.Length);
            Assert.IsType<PdfDictionary>(pages[pageNumber - 1].Value).Items["UserUnit"] = new PdfNumber(userUnit);
            return security.InfoObjectNumber;
        });

    private static byte[] CreateNestedFormImageWithGraphicsStatePdf(
        string outerGraphicsStateEntries,
        string innerGraphicsStateEntries) {
        byte[] pageContent = System.Text.Encoding.ASCII.GetBytes("q /GS1 gs /Fm1 Do Q\n");
        byte[] formContent = System.Text.Encoding.ASCII.GetBytes("q /GS2 gs 80 0 0 40 20 30 cm /Im1 Do Q\n");
        byte[] imageBytes = { 255, 0, 0 };
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 160 160] /Resources << /XObject << /Fm1 5 0 R >> /ExtGState << /GS1 7 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + pageContent.Length + " >>\nstream\n");
        output.Write(pageContent, 0, pageContent.Length);
        WriteAscii(output, "endstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 160 160] /Resources << /XObject << /Im1 6 0 R >> /ExtGState << /GS2 8 0 R >> >> /Length " + formContent.Length + " >>\nstream\n");
        output.Write(formContent, 0, formContent.Length);
        WriteAscii(output, "endstream\nendobj\n");
        WriteAscii(output, "6 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>\nstream\n");
        output.Write(imageBytes, 0, imageBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "7 0 obj\n<< /Type /ExtGState " + outerGraphicsStateEntries + " >>\nendobj\n");
        WriteAscii(output, "8 0 obj\n<< /Type /ExtGState " + innerGraphicsStateEntries + " >>\nendobj\n");
        WriteAscii(output, "trailer\n<< /Root 1 0 R /Size 9 >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] CreateRawImagePdf(
        string content,
        string? graphicsStateEntries = null,
        string? pageEntries = null,
        bool imageMask = false,
        string? secondGraphicsStateEntries = null,
        byte[]? imageBytes = null,
        string? imageDefinition = null,
        string mediaBox = "[0 0 160 160]") {
        byte[] contentBytes = System.Text.Encoding.ASCII.GetBytes(content);
        byte[] resolvedImageBytes = imageBytes ?? (imageMask ? new byte[] { 0x80 } : new byte[] { 255, 0, 0 });
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        string graphicsResources = graphicsStateEntries == null
            ? string.Empty
            : " /ExtGState << /GS1 6 0 R" +
              (secondGraphicsStateEntries == null ? string.Empty : " /GS2 7 0 R") +
              " >>";
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox " + mediaBox + " " + pageEntries + " /Resources << /XObject << /Im1 5 0 R >>" + graphicsResources + " >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "endstream\nendobj\n");
        string resolvedImageDefinition = imageDefinition ?? (imageMask
            ? "/ImageMask true /BitsPerComponent 1"
            : "/ColorSpace /DeviceRGB /BitsPerComponent 8");
        WriteAscii(output, "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 " + resolvedImageDefinition + " /Length " + resolvedImageBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(resolvedImageBytes, 0, resolvedImageBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        if (graphicsStateEntries != null) {
            WriteAscii(output, "6 0 obj\n<< /Type /ExtGState " + graphicsStateEntries + " >>\nendobj\n");
        }
        if (secondGraphicsStateEntries != null) {
            WriteAscii(output, "7 0 obj\n<< /Type /ExtGState " + secondGraphicsStateEntries + " >>\nendobj\n");
        }
        string objectCount = secondGraphicsStateEntries != null ? "8" : graphicsStateEntries == null ? "6" : "7";
        WriteAscii(output, "trailer\n<< /Root 1 0 R /Size " + objectCount + " >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] CreateTextClippedRawImagePdf() {
        const string content = "BT /F1 140 Tf 7 Tr 10 20 Td (M) Tj ET q 40 0 0 40 30 30 cm /Im1 Do Q\n";
        byte[] contentBytes = System.Text.Encoding.ASCII.GetBytes(content);
        byte[] imageBytes = { 255, 0, 0 };
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 160 160] /Resources << /XObject << /Im1 5 0 R >> /Font << /F1 6 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "endstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>\nstream\n");
        output.Write(imageBytes, 0, imageBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "6 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n");
        WriteAscii(output, "trailer\n<< /Root 1 0 R /Size 7 >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] CreateTwoRawImagePdf(string content) {
        byte[] contentBytes = System.Text.Encoding.ASCII.GetBytes(content);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 160 160] /Resources << /XObject << /Im1 5 0 R /Im2 6 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "endstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>\nstream\n");
        output.Write(new byte[] { 255, 0, 0 }, 0, 3);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "6 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>\nstream\n");
        output.Write(new byte[] { 0, 0, 255 }, 0, 3);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "trailer\n<< /Root 1 0 R /Size 7 >>\n%%EOF\n");
        return output.ToArray();
    }

    private static void WriteAscii(Stream output, string value) {
        byte[] bytes = System.Text.Encoding.ASCII.GetBytes(value);
        output.Write(bytes, 0, bytes.Length);
    }
}
