using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlEditableLayoutReviewWave22Tests {
    [Fact]
    public void BorderedImageDescendantKeepsOwningRegionInSemanticFlow() {
        string image = "data:image/png;base64," + Convert.ToBase64String(
            PdfPngTestImages.CreateRgbPng(4, 2));
        string html = "<div style='position:absolute;width:120px;height:60px'>"
            + "<img alt='Bordered picture' src='" + image
            + "' style='width:40px;height:20px;border:3px solid red'></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.NotNull(projection.RemainingDocument.QuerySelector("img[alt='Bordered picture']"));
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "semanticContent=true");
    }

    [Theory]
    [InlineData("hidden")]
    [InlineData("collapse")]
    public void PaintHiddenRegionStaysOutOfNativeProjection(string visibility) {
        string html = "<div style='position:absolute;width:120px;height:40px;visibility:"
            + visibility + ";background:red'>Hidden paint</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Contains("Hidden paint", projection.RemainingDocument.Body!.TextContent,
            StringComparison.Ordinal);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "paintVisible=false; semanticFlow=true");
    }

    [Fact]
    public void WordValidatesOriginalDomAgainstAdapterLimitsBeforeProjection() {
        string children = string.Concat(Enumerable.Range(0, 20)
            .Select(index => "<span>" + index + "</span>"));
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<div style='position:absolute;width:180px;height:50px'>" + children + "</div>",
            new HtmlConversionDocumentOptions { Trust = HtmlInputTrust.Trusted });
        HtmlToWordOptions options = HtmlToWordOptions.CreateTrustedDocumentProfile();
        options.MaxHtmlNodes = 12;

        HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(
            () => source.ToWordDocumentResult(options));

        Assert.Equal(HtmlRenderDiagnosticCodes.NodeLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlConversionLimits.MaxHtmlNodes), exception.LimitSource);
        Assert.Equal(12, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void WordValidatesOriginalDepthAgainstAdapterLimitsBeforeProjection() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<div style='position:absolute;width:180px;height:50px'>"
            + "<span><span><span><span>Deep</span></span></span></span></div>",
            new HtmlConversionDocumentOptions { Trust = HtmlInputTrust.Trusted });
        HtmlToWordOptions options = HtmlToWordOptions.CreateTrustedDocumentProfile();
        options.MaxHtmlDepth = 5;

        HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(
            () => source.ToWordDocumentResult(options));

        Assert.Equal(HtmlConversionDiagnosticCodes.HtmlDepthLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlConversionLimits.MaxHtmlDepth), exception.LimitSource);
        Assert.Equal(5, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void WordProjectsUnderAdapterIntersectedCssComplexityLimits() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<style>.bounded { position:absolute; width:180px; height:50px }</style>"
            + "<div class='bounded'>Bounded</div>",
            new HtmlConversionDocumentOptions { Trust = HtmlInputTrust.Trusted });
        HtmlToWordOptions options = HtmlToWordOptions.CreateTrustedDocumentProfile();
        options.Limits.MaxCssDeclarations = 1;

        HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(
            () => source.ToWordDocumentResult(options));

        Assert.Equal(HtmlConversionDiagnosticCodes.CssDeclarationLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlConversionLimits.MaxCssDeclarations), exception.LimitSource);
        Assert.Equal(1, exception.Limit);
        Assert.True(exception.Actual > exception.Limit);
    }

    [Fact]
    public void ReusedWordRegionImageDoesNotInheritEarlierCropOrTransparency() {
        string image = "data:image/png;base64," + Convert.ToBase64String(
            PdfPngTestImages.CreateRgbPng(4, 2));
        string html = "<div style='position:absolute;left:0;top:0;width:100px;height:60px'>"
            + "<img alt='Affected' src='" + image
            + "' style='width:40px;height:40px;object-fit:cover;object-position:right center;opacity:.4'></div>"
            + "<div style='position:absolute;left:140px;top:0;width:100px;height:60px'>"
            + "<img alt='Clean' src='" + image
            + "' style='width:40px;height:20px;object-fit:fill;opacity:1'></div>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using WordDocument document = result.Value;
        List<WordImage> images = document.TextBoxes
            .SelectMany(textBox => textBox.Paragraphs)
            .Select(paragraph => paragraph.Image)
            .Where(imageItem => imageItem != null)
            .Cast<WordImage>()
            .ToList();
        WordImage affected = Assert.Single(images, imageItem => imageItem.Description == "Affected");
        WordImage clean = Assert.Single(images, imageItem => imageItem.Description == "Clean");

        Assert.Equal(60, affected.Transparency);
        Assert.True(affected.CropLeft.HasValue || affected.CropRight.HasValue
            || affected.CropTop.HasValue || affected.CropBottom.HasValue);
        Assert.Null(clean.Transparency);
        Assert.Null(clean.CropLeft);
        Assert.Null(clean.CropTop);
        Assert.Null(clean.CropRight);
        Assert.Null(clean.CropBottom);
    }
}
