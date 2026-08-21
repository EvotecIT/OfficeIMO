using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Tests.Pdf;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlEditableLayoutWordTests {
    [Fact]
    public void PositionedFormControlsStayInSemanticFlow() {
        const string html = "<div style='position:absolute;width:180px;height:50px'>" +
            "<select name='status'><option>Draft</option><option selected>Approved</option></select></div>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using WordDocument word = result.Value;

        Assert.Empty(word.TextBoxes);
        Assert.Equal("Approved", Assert.Single(word.DropDownLists).SelectedValue);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified);
    }

    [Fact]
    public void PositionedRegionRetainsItsForegroundPictureAndEffects() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(4, 3));
        string html = "<div style='position:absolute;width:180px;height:70px;background-image:url(\"" + image + "\")'>" +
            "<img alt='Hidden marker' src='" + image + "' style='display:none'>" +
            "<img alt='Region marker' src='" + image + "' style='width:24px;height:18px;opacity:.4'></div>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionImageOmitted);
        using var stream = new MemoryStream();
        result.Value.Save(stream);
        result.Value.Dispose();

        byte[] bytes = stream.ToArray();
        using WordDocument reopened = WordDocument.Load(new MemoryStream(bytes),
            new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
        Assert.Single(reopened.TextBoxes);
        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        Assert.Single(package.MainDocumentPart!.ImageParts);
        using var reader = new StreamReader(package.MainDocumentPart!.GetStream());
        string documentXml = reader.ReadToEnd();
        Assert.Contains(":blip", documentXml, StringComparison.Ordinal);
        Assert.Contains(":alphaModFix amt=\"40000\"", documentXml, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.BackgroundLayersFlattened);
    }

    [Fact]
    public async Task RemoteRegionPictureIsFetchedAndEmbeddedByAsyncImport() {
        byte[] png = PdfPngTestImages.CreateRgbPng(4, 3);
        using var httpClient = new HttpClient(new RegionImageHandler(_ => {
            var response = new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new ByteArrayContent(png)
            };
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("image/png");
            return Task.FromResult(response);
        }));
        var options = new HtmlToWordOptions {
            HttpClient = httpClient,
            ImageProcessing = ImageProcessingMode.Embed
        };
        const string html = "<div style='position:absolute;width:180px;height:70px'>" +
            "<img alt='Remote region' src='https://images.example.test/region.png'></div>";

        HtmlToWordResult result = await HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
        using var stream = new MemoryStream();
        result.Value.Save(stream);
        result.Value.Dispose();

        using WordprocessingDocument package = WordprocessingDocument.Open(new MemoryStream(stream.ToArray()), false);
        Assert.Single(package.MainDocumentPart!.ImageParts);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionImageOmitted);
    }

    [Fact]
    public void OversizedPageCoordinateIsBoundedWithStableDiagnostic() {
        const string html = "<div style='position:absolute;left:300000px;top:24px;width:180px;height:70px'>Bounded anchor</div>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using WordDocument word = result.Value;

        Assert.Equal(int.MaxValue, Assert.Single(word.TextBoxes).HorizontalPositionOffset);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("nativeRange", StringComparison.Ordinal));
    }

    [Fact]
    public void OversizedAnchorSizeIsBoundedWithStableDiagnostic() {
        const string html = "<div style='position:absolute;width:1000000000000000px;height:1000000000000000px'>Bounded size</div>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using WordDocument word = result.Value;

        WordTextBox textBox = Assert.Single(word.TextBoxes);
        Assert.Equal(long.MaxValue, textBox.Width);
        Assert.Equal(long.MaxValue, textBox.Height);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("sizeRange", StringComparison.Ordinal));
    }

    [Fact]
    public void MixedInlinePictureRegionStaysOutOfNativeTextBoxes() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<div style='position:absolute;width:180px;height:50px'>Before" +
            "<img alt='Middle' src='" + image + "'>After</div>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using WordDocument word = result.Value;

        Assert.Empty(word.TextBoxes);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Detail == "mixedInlinePictures=true");
    }

    [Fact]
    public void PrintRegionsStaySemanticWhenRenderedPageOwnershipCannotBeMapped() {
        const string html = "<div style='position:absolute;width:160px;height:40px'>Print anchor</div>" +
            "<section style='break-before:page'><p>Later page</p></section>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html, new HtmlConversionDocumentOptions {
            Profile = HtmlConversionProfile.HighFidelityPrint
        });

        HtmlToWordResult result = document.ToWordDocumentResult();
        using WordDocument word = result.Value;

        Assert.Empty(word.TextBoxes);
        Assert.NotEmpty(word.Find("Print anchor", StringComparison.Ordinal));
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified);
    }

    [Fact]
    public void MultiPageContinuousRegionsStaySemanticWithoutAmbiguousPageAnchors() {
        const string html = "<div style='height:1400px'>Long semantic flow</div>" +
            "<div style='position:absolute;top:24px;width:160px;height:40px'>Ambiguous anchor</div>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using WordDocument word = result.Value;

        Assert.Empty(word.TextBoxes);
        Assert.NotEmpty(word.Find("Ambiguous anchor", StringComparison.Ordinal));
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("continuousSurfaceHeight", StringComparison.Ordinal));
    }

    [Fact]
    public void CallerStylesheetContentsParticipateInEditableLayoutProjection() {
        const string html = "<div class='positioned'>Caller styled anchor</div>";
        var options = new HtmlToWordOptions();
        options.StylesheetContents.Add(
            ".positioned{position:absolute;left:24px;top:18px;width:160px;height:40px}");

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
        using WordDocument word = result.Value;

        Assert.Single(word.TextBoxes);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionProjected);
    }

    [Fact]
    public void ConfiguredStylesheetPathsKeepPotentialRegionsInDiagnosedSemanticFlow() {
        string path = Path.GetTempFileName();
        try {
            File.WriteAllText(path,
                ".positioned{position:absolute;left:24px;top:18px;width:160px;height:40px}");
            var options = new HtmlToWordOptions();
            options.StylesheetPaths.Add(path);

            HtmlToWordResult result = HtmlConversionDocument.Parse(
                "<div class='positioned'>Path styled anchor</div>").ToWordDocumentResult(options);
            using WordDocument word = result.Value;

            Assert.Empty(word.TextBoxes);
            Assert.NotEmpty(word.Find("Path styled anchor", StringComparison.Ordinal));
            Assert.Contains(result.Report.Diagnostics, diagnostic =>
                diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
                && diagnostic.Detail == "externalStylesheetSources=true; semanticFlow=true");
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task AllowedDocumentStylesheetLinksKeepPotentialRegionsInDiagnosedSemanticFlow() {
        using var httpClient = new HttpClient(new RegionImageHandler(_ => {
            var response = new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new StringContent(
                    ".positioned{position:absolute;left:24px;top:18px;width:160px;height:40px}")
            };
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/css");
            return Task.FromResult(response);
        }));
        var options = new HtmlToWordOptions {
            AllowDocumentStylesheetLinks = true,
            HttpClient = httpClient
        };
        options.AllowedStylesheetHosts.Add("styles.example.test");
        const string html = "<link rel='stylesheet' href='https://styles.example.test/layout.css'>" +
            "<div class='positioned'>Linked styled anchor</div>";

        HtmlToWordResult result = await HtmlConversionDocument.Parse(html)
            .ToWordDocumentResultAsync(options);
        using WordDocument word = result.Value;

        Assert.Empty(word.TextBoxes);
        Assert.NotEmpty(word.Find("Linked styled anchor", StringComparison.Ordinal));
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "externalStylesheetSources=true; semanticFlow=true");
    }

    [Fact]
    public void PositionedAndFloatingRegionsReopenAsEditableWordAnchors() {
        const string html = "<style>" +
            ".positioned{position:absolute;left:32px;top:24px;width:240px;height:72px;background:#dbeafe;z-index:4}" +
            ".floating{float:right;width:120px;height:48px;background:#fef3c7}" +
            ".flex{display:flex;width:300px}</style>" +
            "<p>Ordinary flow</p><div class='positioned'>Editable positioned</div>" +
            "<div class='floating'>Editable float</div><div class='flex'><span>Flex remains</span></div>";
        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using var stream = new MemoryStream();
        result.Value.Save(stream);
        result.Value.Dispose();

        using WordDocument reopened = WordDocument.Load(
            new MemoryStream(stream.ToArray()),
            new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
        WordTextBox positioned = Assert.Single(reopened.TextBoxes, textBox =>
            textBox.Paragraphs.Any(paragraph => paragraph.Text.Contains("Editable positioned", StringComparison.Ordinal)));
        WordTextBox floating = Assert.Single(reopened.TextBoxes, textBox =>
            textBox.Paragraphs.Any(paragraph => paragraph.Text.Contains("Editable float", StringComparison.Ordinal)));

        Assert.Equal(762000, positioned.HorizontalPositionOffset);
        Assert.Equal(685800, positioned.VerticalPositionOffset);
        Assert.Equal(2286000L, positioned.Width);
        Assert.Equal(685800L, positioned.Height);
        Assert.Equal("DBEAFE", positioned.FillColorHex);
        Assert.Equal(WordImageTextWrapping.InFrontOfText, positioned.WrapText);
        Assert.Equal(WordImageTextWrapping.Square, floating.WrapText);
        Assert.NotEmpty(reopened.Find("Flex remains", StringComparison.Ordinal));
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionProjected);
        Assert.True(result.Succeeded);
    }

    private sealed class RegionImageHandler : HttpMessageHandler {
        private readonly Func<HttpRequestMessage, Task<HttpResponseMessage>> _handler;

        internal RegionImageHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler) {
            _handler = handler;
        }

        protected override Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request,
            CancellationToken cancellationToken) => _handler(request);
    }
}
