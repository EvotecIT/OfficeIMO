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
        WordHtmlImportOptions options = WordHtmlImportOptions.CreateTrustedDocumentProfile();
        options.MaxHtmlNodes = 12;

        HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(
            () => source.ConvertToWordWithEvidence(options));

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
        WordHtmlImportOptions options = WordHtmlImportOptions.CreateTrustedDocumentProfile();
        options.MaxHtmlDepth = 5;

        HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(
            () => source.ConvertToWordWithEvidence(options));

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
        WordHtmlImportOptions options = WordHtmlImportOptions.CreateTrustedDocumentProfile();
        options.Limits.MaxCssDeclarations = 1;

        HtmlConversionLimitException exception = Assert.Throws<HtmlConversionLimitException>(
            () => source.ConvertToWordWithEvidence(options));

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

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ConvertToWordWithEvidence();
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
        Assert.True((affected.CropLeftRatio is double cropRatio2 ? (int?)System.Math.Round(cropRatio2 * 100000D, System.MidpointRounding.AwayFromZero) : null).HasValue || (affected.CropRightRatio is double cropRatio4 ? (int?)System.Math.Round(cropRatio4 * 100000D, System.MidpointRounding.AwayFromZero) : null).HasValue
            || (affected.CropTopRatio is double cropRatio6 ? (int?)System.Math.Round(cropRatio6 * 100000D, System.MidpointRounding.AwayFromZero) : null).HasValue || (affected.CropBottomRatio is double cropRatio8 ? (int?)System.Math.Round(cropRatio8 * 100000D, System.MidpointRounding.AwayFromZero) : null).HasValue);
        Assert.Null(clean.Transparency);
        Assert.Null((clean.CropLeftRatio is double cropRatio10 ? (int?)System.Math.Round(cropRatio10 * 100000D, System.MidpointRounding.AwayFromZero) : null));
        Assert.Null((clean.CropTopRatio is double cropRatio12 ? (int?)System.Math.Round(cropRatio12 * 100000D, System.MidpointRounding.AwayFromZero) : null));
        Assert.Null((clean.CropRightRatio is double cropRatio14 ? (int?)System.Math.Round(cropRatio14 * 100000D, System.MidpointRounding.AwayFromZero) : null));
        Assert.Null((clean.CropBottomRatio is double cropRatio16 ? (int?)System.Math.Round(cropRatio16 * 100000D, System.MidpointRounding.AwayFromZero) : null));
    }
}
