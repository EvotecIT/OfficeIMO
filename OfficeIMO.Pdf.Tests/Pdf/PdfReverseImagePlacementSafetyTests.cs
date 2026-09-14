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

    [Theory]
    [InlineData("/TR << /FunctionType 2 /Domain [0 1] /C0 [0] /C1 [0] /N 1 >>")]
    [InlineData("/op true /OPM 1")]
    [InlineData("/AIS true")]
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
            "/op true /AIS true",
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
        }
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

    private static byte[] CreateImageWithGraphicsStatePdf(string graphicsStateEntries) =>
        CreateRawImagePdf(
            "q /GS1 gs 80 0 0 40 20 30 cm /Im1 Do Q\n",
            graphicsStateEntries);

    private static byte[] CreateRawImagePdf(
        string content,
        string? graphicsStateEntries = null,
        string? pageEntries = null,
        bool imageMask = false,
        string? secondGraphicsStateEntries = null) {
        byte[] contentBytes = System.Text.Encoding.ASCII.GetBytes(content);
        byte[] imageBytes = imageMask ? new byte[] { 0x80 } : new byte[] { 255, 0, 0 };
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.7\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n");
        string graphicsResources = graphicsStateEntries == null
            ? string.Empty
            : " /ExtGState << /GS1 6 0 R" +
              (secondGraphicsStateEntries == null ? string.Empty : " /GS2 7 0 R") +
              " >>";
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 160 160] " + pageEntries + " /Resources << /XObject << /Im1 5 0 R >>" + graphicsResources + " >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "endstream\nendobj\n");
        string imageDefinition = imageMask
            ? "/ImageMask true /BitsPerComponent 1"
            : "/ColorSpace /DeviceRGB /BitsPerComponent 8";
        WriteAscii(output, "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 " + imageDefinition + " /Length " + imageBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(imageBytes, 0, imageBytes.Length);
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

    private static void WriteAscii(Stream output, string value) {
        byte[] bytes = System.Text.Encoding.ASCII.GetBytes(value);
        output.Write(bytes, 0, bytes.Length);
    }
}
