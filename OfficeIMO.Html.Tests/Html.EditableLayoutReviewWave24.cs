using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Rtf;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlEditableLayoutReviewWave24Tests {
    [Fact]
    public void WordRegionBeforeLaterForcedPageBreakStaysInSemanticFlow() {
        const string html = "<p>First page</p>"
            + "<div style='position:absolute;left:24px;top:20px;width:180px;height:70px'>First-page region</div>"
            + "<section style='break-before:page'><p>Later page</p></section>";

        HtmlToWordResult result = HtmlConversionDocument.Parse(html).ToWordDocumentResult();
        using WordDocument document = result.Value;

        Assert.Empty(document.TextBoxes);
        Assert.NotEmpty(document.Find("First-page region", StringComparison.Ordinal));
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "forcedPageBreakAfter=true; semanticFlow=true");
    }

    [Fact]
    public void RtfRegionAfterEarlierForcedPageBreakStaysInSemanticFlow() {
        const string html = "<p>First page</p>"
            + "<section style='break-before:page'>"
            + "<div style='position:absolute;left:24px;top:20px;width:180px;height:70px'>Later-page region</div>"
            + "</section>";

        HtmlToRtfResult result = HtmlConversionDocument.Parse(html).ToRtfDocumentResult();
        IReadOnlyList<IRtfBlock> blocks = result.Value.Blocks;
        int pageBreakIndex = FindParagraphIndex(blocks, paragraph => paragraph.PageBreakBefore);
        int regionIndex = FindParagraphIndex(blocks, paragraph =>
            paragraph.ToPlainText().Contains("Later-page region", StringComparison.Ordinal));
        RtfParagraph region = Assert.IsType<RtfParagraph>(blocks[regionIndex]);

        Assert.True(pageBreakIndex >= 0);
        Assert.True(regionIndex > pageBreakIndex);
        Assert.False(region.Frame.HasAnyValue);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "forcedPageBreakBefore=true; semanticFlow=true");
    }

    [Fact]
    public void PrivateRegionIdentityDoesNotMatchAuthorCssSelectors() {
        const string html = "<style>[data-officeimo-editable-layout-region]{background:#ff0000}</style>"
            + "<div style='position:absolute;width:120px;height:40px'>Transparent region</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));
        HtmlRenderLayoutRegion region = Assert.Single(projection.Regions);

        Assert.Null(region.BackgroundColor);
    }

    [Fact]
    public void AuthoredRegionMarkerStillParticipatesInCssMatching() {
        const string html = "<style>[data-officeimo-editable-layout-region]{padding:12px}</style>"
            + "<div data-officeimo-editable-layout-region='authored' "
            + "style='position:absolute;width:120px;height:40px'>Authored region</div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Equal("authored", projection.RemainingDocument.QuerySelector("div")!
            .GetAttribute("data-officeimo-editable-layout-region"));
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.EffectUnsupported
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("padding", StringComparison.Ordinal));
    }

    [Fact]
    public void AuthoredImageMarkerStillParticipatesInCssMatching() {
        string image = "data:image/png;base64," + Convert.ToBase64String(
            PdfPngTestImages.CreateRgbPng(4, 2));
        string html = "<style>[data-officeimo-editable-layout-image]{border:3px solid red}</style>"
            + "<div style='position:absolute;width:120px;height:60px'>"
            + "<img data-officeimo-editable-layout-image='authored' alt='Bordered' src='" + image + "'></div>";

        HtmlEditableLayoutProjection projection = HtmlEditableLayoutProjector.Project(
            HtmlConversionDocument.Parse(html));

        Assert.Empty(projection.Regions);
        Assert.Equal("authored", projection.RemainingDocument.QuerySelector("img")!
            .GetAttribute("data-officeimo-editable-layout-image"));
    }

    [Fact]
    public void PowerPointRetriesShapeReservationAfterInFlowPlacementRollback() {
        const string html = "<h1>Title</h1>"
            + "<div style='display:grid;width:40px;height:60px'>Rejected flow</div>"
            + "<div style='position:absolute;left:20px;top:20px;width:20px;height:20px'>Retried positioned</div>";
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxShapes = 2;
        limits.MaxAbsoluteGeometry = 50D;

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions {
                Mode = HtmlImportMode.Generic,
                Limits = limits
            });
        using PowerPointPresentation presentation = result.Value;
        PowerPointSlide slide = Assert.Single(presentation.Slides);

        Assert.DoesNotContain(slide.TextBoxes, box => box.Text == "Rejected flow");
        Assert.Contains(slide.TextBoxes, box => box.Text == "Retriedpositioned");
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded
            && diagnostic.Message.Contains("native shape limit", StringComparison.Ordinal));
    }

    [Fact]
    public async Task RemoteWordRegionImageStaysSemanticUntilAsyncResourceResolution() {
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
        const string html = "<div style='position:absolute;width:180px;height:70px'>"
            + "<img alt='Remote region' src='https://images.example.test/region.png'></div>";

        HtmlToWordResult result = await HtmlConversionDocument.Parse(html)
            .ToWordDocumentResultAsync(options);
        using WordDocument document = result.Value;

        Assert.Empty(document.TextBoxes);
        WordImage image = Assert.Single(document.Images);
        Assert.Equal("Remote region", image.Description);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail == "unrenderedRegionImage=true; semanticFlow=true");
    }

    private static int FindParagraphIndex(
        IReadOnlyList<IRtfBlock> blocks,
        Func<RtfParagraph, bool> predicate) {
        for (int index = 0; index < blocks.Count; index++) {
            if (blocks[index] is RtfParagraph paragraph && predicate(paragraph)) return index;
        }
        return -1;
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