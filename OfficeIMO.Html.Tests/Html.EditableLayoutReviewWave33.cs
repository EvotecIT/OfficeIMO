using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlEditableLayoutReviewWave33Tests {
    [Theory]
    [InlineData(false, "hidden")]
    [InlineData(false, "collapse")]
    [InlineData(true, "hidden")]
    [InlineData(true, "collapse")]
    public void VisibleImageInsideHiddenAncestorKeepsItsRenderedSourceAssociation(
        bool useStylesheet,
        string ancestorVisibility) {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(4, 3));
        string visibility = useStylesheet
            ? "<style>.hidden-parent{visibility:" + ancestorVisibility + "}.visible-image{visibility:visible}</style>"
                + "<span class='hidden-parent'><img class='visible-image' alt='Visible override' src='" + image + "'></span>"
            : "<span style='visibility:" + ancestorVisibility + "'><img style='visibility:visible' alt='Visible override' src='"
                + image + "'></span>";
        string html = "<div style='position:absolute;width:180px;height:70px'>" + visibility + "</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        HtmlRenderLayoutRegion region = Assert.Single(projection.Regions);
        AngleSharp.Html.Dom.IHtmlImageElement source = Assert.Single(projection.GetSourceImages(region));
        HtmlRenderImage rendered = Assert.Single(HtmlEditableLayoutProjector
            .EnumerateImages(region.Visuals, includeBackgroundImages: false)
            .Select(item => item.Image));
        Assert.Equal("Visible override", source.AlternativeText);
        Assert.Same(source, projection.GetSourceImage(rendered));
    }

    [Fact]
    public void DisplayNoneAncestorStillPrunesExplicitlyVisibleImage() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(4, 3));
        string html = "<div style='position:absolute;width:180px;height:70px'>"
            + "<span style='display:none'><img style='visibility:visible' alt='Pruned' src='" + image + "'></span>"
            + "</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        HtmlRenderLayoutRegion region = Assert.Single(projection.Regions);
        Assert.Empty(projection.GetSourceImages(region));
        Assert.Empty(HtmlEditableLayoutProjector.EnumerateImages(region.Visuals, includeBackgroundImages: false));
    }

    [Fact]
    public void WordEmbedsVisibleImageInsideHiddenAncestor() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(4, 3));
        string html = "<div style='position:absolute;width:180px;height:70px'>"
            + "<span style='visibility:hidden'><img style='visibility:visible' alt='Word override' src='" + image + "'></span>"
            + "</div>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using var stream = new MemoryStream();
        result.Value.Save(stream);
        result.Value.Dispose();

        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(stream.ToArray()), false);
        Assert.Single(package.MainDocumentPart!.ImageParts);
    }

    [Fact]
    public void RtfEmbedsVisibleImageInsideHiddenAncestor() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(4, 3));
        string html = "<div style='position:absolute;width:180px;height:70px'>"
            + "<span style='visibility:hidden'><img style='visibility:visible' alt='RTF override' src='" + image + "'></span>"
            + "</div>";

        HtmlToRtfResult result = HtmlConversionDocument.Parse(html).ToRtfDocumentResult();

        RtfParagraph paragraph = Assert.Single(result.Value.Paragraphs, item => item.Inlines.OfType<RtfImage>().Any());
        Assert.Single(paragraph.Inlines.OfType<RtfImage>());
    }
}