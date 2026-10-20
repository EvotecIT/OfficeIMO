using OfficeIMO.Drawing;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Mhtml;
using OfficeIMO.Tests.Pdf;
using System.Threading.Tasks;
using System.Xml.Linq;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class MhtmlRenderResourcePolicyTests {
    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged)]
    [InlineData(HtmlRenderIntentProfile.ScreenSnapshotPaged)]
    public async Task PdfCapableImageRequestsRetainArchiveStylesAndMatchDirectPdf(HtmlRenderIntentProfile profile) {
        var document = new MhtmlDocument(
            "<link rel='stylesheet' href='cid:style'><p>Body</p><img src='cid:logo' width='8' height='8'>",
            new[] {
                new MhtmlResource(Encoding.UTF8.GetBytes(
                    "body{margin:0;background:#d0ebf8}p{margin:0;color:#336699}p::before{content:'Archived CSS '}"),
                    "text/css", contentId: "style"),
                new MhtmlResource(PdfPngTestImages.CreateRgbPng(2, 2), "image/png", contentId: "logo")
            }, "https://example.test/page.html");
        var options = CreateOptions();
        HtmlRenderRequest pdfRequest = HtmlRenderRequest.Create(profile, HtmlRenderEncoder.Pdf, options);
        PdfCore.PdfDocumentConversionResult pdf = await document.RenderToPdfDocumentResultAsync(pdfRequest);
        var expected = PdfCore.PdfImageExportExtensions.ExportImages(pdf, OfficeImageExportFormat.Png,
            new PdfCore.PdfImageExportOptions { Scale = 96D / 72D });
        Assert.Contains("Archived CSS", PdfCore.PdfReadDocument.Open(pdf.ToBytes()).ExtractText(), StringComparison.Ordinal);

        document.ConfigureRenderOptions(options);
        HtmlRenderRequest imageRequest = pdfRequest.WithOptions(options).WithEncoder(HtmlRenderEncoder.Png);
        HtmlRenderResult display = await HtmlRenderEngine.ExecuteAsync(document.HtmlDocument,
            imageRequest.WithEncoder(HtmlRenderEncoder.DisplayList));
        Assert.Single(display.Document.Pages.SelectMany(page => Walk(page.Scene)).OfType<HtmlRenderImage>());
        Assert.Contains("Archived CSS", string.Concat(display.Document.Pages.SelectMany(page => Walk(page.Scene))
            .OfType<HtmlRenderText>().Select(text => text.Text)), StringComparison.Ordinal);
        HtmlRenderResult rendered = await HtmlRenderEngine.ExecuteAsync(document.HtmlDocument, imageRequest);
        var actual = rendered.ExportImages();
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable);
        Assert.Equal(expected.Count, actual.Count);
        for (int index = 0; index < expected.Count; index++) Assert.Equal(expected[index].Bytes, actual[index].Bytes);

        HtmlRenderResult svg = await HtmlRenderEngine.ExecuteAsync(document.HtmlDocument,
            imageRequest.WithEncoder(HtmlRenderEncoder.Svg));
        string[] vectorPages = svg.ExportImages().Select(page => Encoding.UTF8.GetString(page.Bytes)).ToArray();
        string vectors = string.Concat(vectorPages);
        string vectorText = string.Concat(vectorPages.SelectMany(page =>
            XDocument.Parse(page).Descendants(XName.Get("text", "http://www.w3.org/2000/svg")))
            .Select(text => text.Value));
        Assert.Contains("Archived CSS", vectorText, StringComparison.Ordinal);
        Assert.Contains("#336699", vectors, StringComparison.OrdinalIgnoreCase);
        Assert.Single(svg.Document.Pages.SelectMany(page => Walk(page.Scene)).OfType<HtmlRenderImage>());
        Assert.False(options.ResourcePolicy.AllowRemoteResourceResolution);
        Assert.False(options.ResourcePolicy.AllowLocalFileAccess);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task ReconfiguredPdfImageRequestsKeepEmbeddedAndHostPermissionsIndependent(
        bool allowEmbedded, bool allowRemote) {
        byte[] image = PdfPngTestImages.CreateRgbPng(2, 2);
        var document = new MhtmlDocument(
            "<img src='cid:logo' width='8' height='8'>"
            + "<img src='https://example.test/allowed.png' width='8' height='8'>"
            + "<img src='https://example.test/blocked.png' width='8' height='8'>"
            + "<img src='file:///outside/missing.png' width='8' height='8'>",
            new[] { new MhtmlResource(image, "image/png", contentId: "logo") },
            "https://example.test/page.html");
        var options = CreateOptions();
        options.ResourcePolicy.AllowRemoteResourceResolution = allowRemote;
        options.ResourceUrlPolicy = HtmlUrlPolicy.CreateWebResourceProfile();
        options.ResourceUrlPolicy.ResolvedUrlTransform = uri => uri.EndsWith("blocked.png", StringComparison.Ordinal)
            ? null : uri;
        int unwrappedCalls = 0;
        options.ResourceResolver = (request, token) => {
            unwrappedCalls++;
            return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(image, "image/png"));
        };
        var fetched = new List<Uri>();
        MhtmlRemoteResourcePolicy remote = MhtmlRemoteResourcePolicy.CreateSameOriginProfile();
        remote.ResourceFetcher = (request, token) => {
            fetched.Add(request.Uri);
            return Task.FromResult<MhtmlRemoteResourceResponse?>(new MhtmlRemoteResourceResponse(image, "image/png"));
        };
        options.ResourcePolicy.AllowEmbeddedPackageResources = allowEmbedded;
        document.ConfigureRenderOptions(options, remote);
        options.ResourcePolicy.AllowEmbeddedPackageResources = true;
        document.ConfigureRenderOptions(options, remote);
        document.ConfigureRenderOptions(options, remote);
        // Tightening after setup must also close the general host resolver's archive alias.
        options.ResourcePolicy.AllowEmbeddedPackageResources = allowEmbedded;

        HtmlRenderResult rendered = await HtmlRenderEngine.ExecuteAsync(document.HtmlDocument,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Png, options));
        int expectedImages = (allowEmbedded ? 1 : 0) + (allowRemote ? 1 : 0);
        Assert.Equal(expectedImages, rendered.Document.Pages.SelectMany(page => Walk(page.Scene)).OfType<HtmlRenderImage>().Count());
        Assert.Equal(allowRemote ? 1 : 0, fetched.Count);
        Assert.All(fetched, uri => Assert.Equal("https://example.test/allowed.png", uri.AbsoluteUri));
        Assert.Equal(0, unwrappedCalls);
        Assert.DoesNotContain("cid", options.UrlPolicy.AllowedUrlSchemes);
        Assert.True(options.UrlPolicy.DisallowFileUrls);
        Assert.False(options.ResourcePolicy.AllowLocalFileAccess);

        // A later explicit host delegate must not be silently replaced by the stored container resolver.
        options.ResourceResolver = (request, token) => {
            unwrappedCalls++;
            return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(image, "image/png"));
        };
        HtmlRenderResult replaced = await HtmlRenderEngine.ExecuteAsync(document.HtmlDocument,
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Png, options));
        Assert.Equal(expectedImages, replaced.Document.Pages.SelectMany(page => Walk(page.Scene)).OfType<HtmlRenderImage>().Count());
        Assert.Equal(allowRemote ? 1 : 0, unwrappedCalls);
    }

    private static HtmlToPdfOptions CreateOptions() => new() {
        PageSize = new OfficePageSize(160D / 96D, 120D / 96D),
        Margins = HtmlRenderMargins.All(8D),
        ViewportWidth = 160D,
        ViewportHeight = 120D,
        ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
    };

    private static IEnumerable<HtmlRenderVisual> Walk(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderClipGroup group => group.Visuals,
                HtmlRenderPathClipGroup group => group.Visuals,
                HtmlRenderEffectGroup group => group.Visuals,
                HtmlRenderSemanticGroup group => group.Visuals,
                HtmlRenderLogicalTextGroup group => group.Visuals,
                HtmlRenderLayoutRegion group => group.Visuals,
                _ => null
            };
            if (children == null) continue;
            foreach (HtmlRenderVisual child in Walk(children)) yield return child;
        }
    }
}
