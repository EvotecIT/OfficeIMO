using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Html;
using OfficeIMO.Mhtml;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class MhtmlEmbeddedImageConversionTests {
    [Fact]
    public void ArchivedImageBecomesAnEditableWordImageWithoutChangingTheArchive() {
        byte[] png = PdfPngTestImages.CreateRgbPng(2, 2);
        var original = new MhtmlDocument(
            "<h1>Gallery</h1><a href='https://example.test/gallery'>"
                + "<div><img alt='Gallery photo' src='images/gallery.png'></div>"
                + "<div>Gallery photo</div></a>",
            new[] { new MhtmlResource(png, "image/png",
                contentLocation: "https://example.test/page/images/gallery.png") },
            contentLocation: "https://example.test/page/index.html");
        using var archiveBytes = new MemoryStream();
        original.Save(archiveBytes);
        archiveBytes.Position = 0;
        MhtmlDocument archive = MhtmlDocument.Load(archiveBytes);

        MhtmlImageEmbeddingResult prepared = archive.CreateEmbeddedImageDocumentResult();

        Assert.True(prepared.Report.Succeeded);
        Assert.False(prepared.Report.HasLoss);
        Assert.Equal(1, prepared.EmbeddedResourceCount);
        Assert.Equal(png.Length, prepared.EmbeddedResourceBytes);
        Assert.Contains("data:image/png;base64,", prepared.Value.SourceHtml, StringComparison.Ordinal);
        Assert.Contains("images/gallery.png", archive.HtmlDocument.SourceHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("data:image/", archive.HtmlDocument.SourceHtml, StringComparison.Ordinal);

        HtmlToWordResult word = prepared.Value.ToWordDocumentResult();
        Assert.True(word.Report.Succeeded);
        using var packageBytes = new MemoryStream();
        using (var document = word.RequireValue()) document.Save(packageBytes);
        using WordprocessingDocument package = WordprocessingDocument.Open(
            new MemoryStream(packageBytes.ToArray()), false);
        Assert.Single(package.MainDocumentPart!.ImageParts);
    }

    [Fact]
    public void MissingAndOversizeImagesStayExternalAndProduceResourceLosses() {
        byte[] png = PdfPngTestImages.CreateRgbPng(2, 2);
        var archive = new MhtmlDocument(
            "<img src='present.png'><img src='missing.png'>",
            new[] { new MhtmlResource(png, "image/png",
                contentLocation: "https://example.test/page/present.png") },
            contentLocation: "https://example.test/page/index.html");

        MhtmlImageEmbeddingResult prepared = archive.CreateEmbeddedImageDocumentResult(
            new HtmlRenderOptions { MaxResourceBytes = 1 });

        Assert.Equal(0, prepared.EmbeddedResourceCount);
        Assert.Equal(0, prepared.EmbeddedResourceBytes);
        Assert.True(prepared.Report.HasLoss);
        Assert.Contains(prepared.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded
                && diagnostic.Source == "present.png");
        Assert.Contains(prepared.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable
                && diagnostic.Source == "missing.png");
        Assert.DoesNotContain("data:image/", prepared.Value.SourceHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void ChromiumContentLocationCidWithoutContentIdStillEmbedsTheImage() {
        byte[] png = PdfPngTestImages.CreateRgbPng(2, 2);
        var archive = new MhtmlDocument(
            "<img src='cid:gallery@archive'>",
            new[] { new MhtmlResource(png, "image/png", contentLocation: "cid:gallery@archive") },
            contentLocation: "https://example.test/index.html");

        MhtmlImageEmbeddingResult prepared = archive.CreateEmbeddedImageDocumentResult();

        Assert.Equal(1, prepared.EmbeddedResourceCount);
        Assert.False(prepared.Report.HasLoss);
        Assert.Contains("data:image/png;base64,", prepared.Value.SourceHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void PictureFallbackImageIsEmbeddedEvenWhenAResponsiveSourceIsSelected() {
        byte[] png = PdfPngTestImages.CreateRgbPng(2, 2);
        var archive = new MhtmlDocument(
            "<picture><source srcset='selected.webp' type='image/webp'>"
                + "<img src='fallback.png' alt='Fallback'></picture>",
            new[] { new MhtmlResource(png, "image/png",
                contentLocation: "https://example.test/page/fallback.png") },
            contentLocation: "https://example.test/page/index.html");

        MhtmlImageEmbeddingResult prepared = archive.CreateEmbeddedImageDocumentResult();

        Assert.Equal(1, prepared.EmbeddedResourceCount);
        Assert.Contains("data:image/png;base64,", prepared.Value.SourceHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void RepeatedImagesStayWithinTheEditableHtmlSourceLimit() {
        byte[] png = PdfPngTestImages.CreateRgbPng(2, 2);
        var options = new HtmlConversionDocumentOptions {
            Limits = new HtmlConversionLimits { MaxInputCharacters = 1200 }
        };
        var archive = new MhtmlDocument(
            string.Concat(Enumerable.Repeat("<img src='repeat.png'>", 20)),
            new[] { new MhtmlResource(png, "image/png",
                contentLocation: "https://example.test/page/repeat.png") },
            contentLocation: "https://example.test/page/index.html",
            htmlOptions: options);

        MhtmlImageEmbeddingResult prepared = archive.CreateEmbeddedImageDocumentResult();

        Assert.True(prepared.Report.HasLoss);
        Assert.Contains(prepared.Report.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.InputCharacterLimitExceeded);
        Assert.Contains("repeat.png", prepared.Value.SourceHtml, StringComparison.Ordinal);
        Assert.True(prepared.Value.SourceHtml.Length <= 1200);
    }
}
