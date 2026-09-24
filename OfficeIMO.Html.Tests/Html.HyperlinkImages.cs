using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Blip = DocumentFormat.OpenXml.Drawing.Blip;
using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlHyperlinkImageTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LinkedImagesRetainTheirMediaPartAndDrawing(bool blockContent) {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string content = blockContent
            ? "<div><img src='" + image + "' alt='Photo'></div><div>Photo title</div>"
            : "<img src='" + image + "' alt='Photo'>Photo title";
        string html = "<p><a href='https://example.test/gallery'>" + content + "</a></p>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using var stream = new MemoryStream();
        using (var document = result.RequireValue()) document.Save(stream);
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(stream.ToArray()), false);

        Assert.True(result.Report.Succeeded);
        Assert.Single(package.MainDocumentPart!.ImageParts);
        Assert.Contains(package.MainDocumentPart.Document.Descendants<Hyperlink>(),
            link => link.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any());
    }

    [Fact]
    public void RepeatedImageInBodyAndHeaderHasRelationshipsInBothStories() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<p><img src='" + image + "'></p>"
            + "<header class='word-header'><a href='https://example.test/gallery'>"
            + "<img src='" + image + "'></a></header>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using var stream = new MemoryStream();
        using (var document = result.RequireValue()) document.Save(stream);
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(stream.ToArray()), false);

        MainDocumentPart main = package.MainDocumentPart!;
        Assert.Single(main.ImageParts);
        HeaderPart header = Assert.Single(main.HeaderParts);
        Assert.Single(header.ImageParts);
        var blip = Assert.Single(header.Header.Descendants<Blip>());
        Assert.IsType<ImagePart>(header.GetPartById(blip.Embed!.Value!));
    }
}
