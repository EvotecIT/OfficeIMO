using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Mhtml;
using OfficeIMO.Rtf;
using OfficeIMO.Tests.Pdf;
using System.Threading.Tasks;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed class HtmlEditableLayoutReviewWave21Tests {
    [Fact]
    public void NegativeRtfFrameCoordinatesUseNegativeCapableControls() {
        const string html = "<div style='position:absolute;left:-200px;top:-100px;width:180px;height:70px'>Negative frame</div>";

        HtmlToRtfResult result = HtmlConversionDocument.Parse(html).ToRtfDocumentResult();
        RtfParagraph paragraph = Assert.Single(result.Value.Paragraphs, item =>
            item.ToPlainText().Contains("Negative frame", StringComparison.Ordinal));
        string rtf = result.Value.ToRtf();

        Assert.Equal(RtfParagraphFrameHorizontalPosition.NegativeAbsolute, paragraph.Frame.HorizontalPosition);
        Assert.Equal(RtfParagraphFrameVerticalPosition.NegativeAbsolute, paragraph.Frame.VerticalPosition);
        Assert.Contains(@"\posnegx-", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\posnegy-", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void FirstPageRtfFrameIsInsertedBeforeLaterForcedPageBreak() {
        const string html = "<p>First page</p>"
            + "<div style='position:absolute;left:24px;top:20px;width:180px;height:70px'>First-page frame</div>"
            + "<section style='break-before:page'><p>Later page</p></section>";

        HtmlToRtfResult result = HtmlConversionDocument.Parse(html).ToRtfDocumentResult();
        IReadOnlyList<IRtfBlock> blocks = result.Value.Blocks;
        int frameIndex = FindParagraphIndex(blocks, paragraph => paragraph.Frame.HasAnyValue);
        int pageBreakIndex = FindParagraphIndex(blocks, paragraph => paragraph.PageBreakBefore);

        Assert.True(frameIndex >= 0);
        Assert.True(pageBreakIndex >= 0);
        Assert.True(frameIndex < pageBreakIndex);
        Assert.Contains(@"\phpg", result.Value.ToRtf(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task MhtmlPdfRejectsUnwrappedLegacyRemoteResolver(bool allowEmbeddedResources) {
        var document = new MhtmlDocument(
            "<img src='https://example.test/legacy.png' width='20' height='20'>",
            contentLocation: "https://example.test/page.html");
        int resolverCalls = 0;
        PdfCore.PdfResourcePolicy resourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost();
        resourcePolicy.AllowEmbeddedPackageResources = allowEmbeddedResources;
        var options = new HtmlToPdfOptions {
            ResourcePolicy = resourcePolicy,
            ResourceResolver = (request, cancellationToken) => {
                resolverCalls++;
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(
                    PdfPngTestImages.CreateRgbPng(2, 2),
                    "image/png",
                    new Uri("https://other.test/legacy.png"),
                    redirectCount: 1));
            }
        };

        PdfCore.PdfDocumentConversionResult result = await document.ToPdfDocumentResultAsync(options);

        Assert.Equal(0, resolverCalls);
        Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(result.ToBytes()).Where(image => image.IsImageFile));
        Assert.Contains(result.Warnings, warning =>
            warning.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable
            || warning.Code == HtmlRenderDiagnosticCodes.ExternalImagePending);
    }

    [Fact]
    public async Task MhtmlPdfRetainsPolicyOwnedRemoteResolverWhenEmbeddedResourcesAreDisabled() {
        var document = new MhtmlDocument(
            "<img src='https://example.test/remote.png' width='20' height='20'>",
            contentLocation: "https://example.test/page.html");
        int fetcherCalls = 0;
        PdfCore.PdfResourcePolicy resourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost();
        resourcePolicy.AllowEmbeddedPackageResources = false;
        var options = new HtmlToPdfOptions { ResourcePolicy = resourcePolicy };
        MhtmlRemoteResourcePolicy remotePolicy = MhtmlRemoteResourcePolicy.CreateSameOriginProfile();
        remotePolicy.ResourceFetcher = (request, cancellationToken) => {
            fetcherCalls++;
            return Task.FromResult<MhtmlRemoteResourceResponse?>(new MhtmlRemoteResourceResponse(
                PdfPngTestImages.CreateRgbPng(2, 2),
                "image/png"));
        };
        document.ConfigureRenderOptions(options, remotePolicy);

        PdfCore.PdfDocumentConversionResult result = await document.ToPdfDocumentResultAsync(options);

        Assert.Equal(1, fetcherCalls);
        Assert.Single(PdfCore.PdfImageExtractor.ExtractImages(result.ToBytes()), image => image.IsImageFile);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable);
    }

    private static int FindParagraphIndex(
        IReadOnlyList<IRtfBlock> blocks,
        Func<RtfParagraph, bool> predicate) {
        for (int index = 0; index < blocks.Count; index++) {
            if (blocks[index] is RtfParagraph paragraph && predicate(paragraph)) return index;
        }
        return -1;
    }
}
