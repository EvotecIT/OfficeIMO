using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Html;
using OfficeIMO.Mhtml;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word.Html;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class MhtmlEmbeddedImageConversionTests {
    [Fact]
    public async Task ClosedDetailsKeepsItsSummaryButNotItsHiddenBody() {
        const string html = "<details><summary>How this works</summary><p>Closed explanation</p></details>"
            + "<details open><summary>More detail</summary><p>Open explanation</p></details>"
            + "<p>Article</p>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);

        HtmlVisibleContentResult visible = await source.CreateVisibleContentDocumentResultAsync();

        Assert.Contains("How this works", visible.Value.SourceHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("Closed explanation", visible.Value.SourceHtml, StringComparison.Ordinal);
        Assert.Contains("Open explanation", visible.Value.SourceHtml, StringComparison.Ordinal);
        Assert.Contains("Article", visible.Value.SourceHtml, StringComparison.Ordinal);
        Assert.Contains("Closed explanation", source.SourceHtml, StringComparison.Ordinal);
        Assert.True(visible.OmittedElementCount > 0);
    }

    [Theory]
    [InlineData("html { opacity: 0 }")]
    [InlineData("body { display: none }")]
    public async Task HiddenDocumentRootDoesNotExportBodyContent(string css) {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<style>" + css + "</style><body><p>Hidden article</p></body>");

        HtmlVisibleContentResult result = await source.CreateVisibleContentDocumentResultAsync();

        Assert.True(result.OmittedElementCount > 0);
        Assert.DoesNotContain("Hidden article", result.Value.SourceHtml, StringComparison.Ordinal);
        Assert.Contains("Hidden article", source.SourceHtml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ArchivedShadowSnapshotIsProjectedForEditableContent() {
        const string html = "<sample-card><template shadowmode='open'><p>Shadow article</p>"
            + "</template><p>Unused light content</p></sample-card>";
        var archive = new MhtmlDocument(html);
        var options = new HtmlRenderOptions();
        archive.ConfigureRenderOptions(options);

        HtmlVisibleContentResult result = await archive.CreateEmbeddedImageDocumentResult(options)
            .RequireValue().CreateVisibleContentDocumentResultAsync(options);

        Assert.Contains("Shadow article", result.Value.SourceHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("Unused light content", result.Value.SourceHtml, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.SerializedShadowRootApproximated);
        Assert.True(result.Report.HasLoss);
        Assert.Contains("Unused light content", archive.Html, StringComparison.Ordinal);
    }

    [Fact]
    public async Task PolicyBlockedStylesheetIsReportedAsVisibilityLoss() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<link rel='stylesheet' href='file:///private/article.css'><p>Article</p>");

        HtmlVisibleContentResult result = await source.CreateVisibleContentDocumentResultAsync();

        Assert.Equal(0, result.AppliedStylesheetCount);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Source == "file:///private/article.css"
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.True(result.Report.HasLoss);
        Assert.Throws<HtmlConversionException>(() => result.RequireNoLoss());
    }

    [Fact]
    public async Task MissingStylesheetResolverIsReportedBeforeEditableImport() {
        HtmlConversionDocument source = HtmlConversionDocument.Parse(
            "<link rel='stylesheet' href='https://example.test/article.css'><p>Article</p>");

        HtmlVisibleContentResult result = await source.CreateVisibleContentDocumentResultAsync();

        Assert.Equal(0, result.AppliedStylesheetCount);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ExternalStylesheetPending);
        Assert.True(result.Report.HasLoss);
        Assert.Contains("Article", result.Value.SourceHtml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ArchivedCssFiltersHiddenEditableContentWithoutImportingTheWholeStylesheet() {
        const string html = "<html><head><link rel='stylesheet' href='styles.css'></head><body>"
            + "<p>Visible article</p><div class='print-hidden'>Print navigation</div>"
            + "<div class='transparent'>Closed menu</div></body></html>";
        const string css = ".print-hidden{display:block}@media print{.print-hidden{display:none}}"
            + ".transparent{opacity:0}";
        var archive = new MhtmlDocument(html,
            new[] { new MhtmlResource(Encoding.UTF8.GetBytes(css), "text/css",
                contentLocation: "https://example.test/page/styles.css") },
            contentLocation: "https://example.test/page/index.html");
        MhtmlImageEmbeddingResult prepared = archive.CreateEmbeddedImageDocumentResult();
        var printOptions = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged };
        archive.ConfigureRenderOptions(printOptions);

        HtmlVisibleContentResult print = await prepared.Value.CreateVisibleContentDocumentResultAsync(printOptions);

        Assert.Equal(1, print.AppliedStylesheetCount);
        Assert.True(print.OmittedElementCount >= 2);
        Assert.Contains("Visible article", print.Value.SourceHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("Print navigation", print.Value.SourceHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("Closed menu", print.Value.SourceHtml, StringComparison.Ordinal);
        Assert.DoesNotContain(css, print.Value.SourceHtml, StringComparison.Ordinal);
        Assert.Contains("styles.css", print.Value.SourceHtml, StringComparison.Ordinal);
        Assert.Contains("Print navigation", archive.Html, StringComparison.Ordinal);

        HtmlToWordResult word = print.Value.ToWordDocumentResult();
        using var output = new MemoryStream();
        using (var document = word.RequireValue()) document.Save(output);
        using WordprocessingDocument saved = WordprocessingDocument.Open(
            new MemoryStream(output.ToArray()), false);
        string savedText = saved.MainDocumentPart!.Document.Body!.InnerText;
        Assert.Contains("Visible article", savedText, StringComparison.Ordinal);
        Assert.DoesNotContain("Print navigation", savedText, StringComparison.Ordinal);
        Assert.DoesNotContain("Closed menu", savedText, StringComparison.Ordinal);

        var screenOptions = new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous };
        archive.ConfigureRenderOptions(screenOptions);
        HtmlVisibleContentResult screen = await prepared.Value.CreateVisibleContentDocumentResultAsync(screenOptions);
        Assert.Contains("Print navigation", screen.Value.SourceHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("Closed menu", screen.Value.SourceHtml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ArchivedImageMaximumWidthIsCarriedIntoEditableWord() {
        byte[] png = PdfPngTestImages.CreateRgbPng(800, 300);
        var archive = new MhtmlDocument(
            "<link rel='stylesheet' href='styles.css'><img class='logo' src='images/logo.png' alt='Site logo'>",
            new[] {
                new MhtmlResource(Encoding.UTF8.GetBytes(".logo{max-width:210px}"), "text/css",
                    contentLocation: "https://example.test/page/styles.css"),
                new MhtmlResource(png, "image/png",
                    contentLocation: "https://example.test/page/images/logo.png")
            },
            contentLocation: "https://example.test/page/index.html");
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged };
        archive.ConfigureRenderOptions(options);

        HtmlConversionDocument images = archive.CreateEmbeddedImageDocumentResult().RequireValue();
        HtmlVisibleContentResult visible = await images.CreateVisibleContentDocumentResultAsync(options);
        HtmlToWordResult converted = visible.Value.ToWordDocumentResult();
        using var document = converted.RequireValue();

        Assert.Equal(1, visible.AppliedStylesheetCount);
        Assert.Contains("max-width:210px", visible.Value.SourceHtml, StringComparison.Ordinal);
        Assert.Equal(210D, Assert.Single(document.Images).Width!.Value, precision: 2);
        Assert.Contains(converted.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.ContentApproximated
            && diagnostic.Source == "Site logo");
        Assert.DoesNotContain("max-width:210px", archive.Html, StringComparison.Ordinal);
    }

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
