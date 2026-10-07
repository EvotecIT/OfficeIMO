using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed class HtmlSiteBundleTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SiteBundleLoadsNestedCssAndImagesAndReopensSvgAndPdfWithoutHostAccess(bool asynchronous) {
        using MemoryStream zip = CreateBundle(
            ("articles/index.html", Utf8("<link rel='stylesheet' href='../css/site.css?v=1'><p>ARCHIVED</p>"
                + "<img src='../images/photo.png#frame' width='8' height='8'>")),
            ("css/site.css", Utf8("@import 'theme.css'; p{color:#123456}")),
            ("css/theme.css", Utf8("html,body{margin:0} p{font-size:20px}")),
            ("images/photo.png", PdfPngTestImages.CreateRgbPng(2, 2)));
        zip.Position = 9;
        HtmlSiteBundle bundle = await HtmlSiteBundle.LoadAsync(zip,
            new HtmlSiteBundleOptions { ArchiveBaseUri = new Uri("https://archive.example.test/site/") });
        Assert.Equal(9, zip.Position);
        Assert.True(zip.CanRead);
        Assert.Equal("articles/index.html", bundle.EntryPath);
        Assert.Equal(new Uri("https://archive.example.test/site/articles/index.html"), bundle.BaseUri);

        var options = new HtmlRenderOptions { Margins = HtmlRenderMargins.All(0) };
        HtmlRenderRequest original = HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Svg, options);
        HtmlRenderResult rendered = asynchronous
            ? await HtmlRenderEngine.ExecuteAsync(bundle.HtmlDocument, bundle.CreateRenderRequest(original))
            : HtmlRenderEngine.Execute(bundle.HtmlDocument, bundle.CreateRenderRequest(original));
        string svg = Encoding.UTF8.GetString(rendered.ExportImage().Bytes);
        XDocument parsedSvg = XDocument.Parse(svg);
        Assert.Single(parsedSvg.Descendants(), element => element.Name.LocalName == "image");
        Assert.Contains("#123456", svg, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable);
        Assert.Null(original.Options.ResourceResolver);
        Assert.Null(original.Options.BaseUri);

        var pdfOptions = new HtmlToPdfOptions { ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic() };
        Assert.True(pdfOptions.ResourcePolicy.AllowEmbeddedPackageResources);
        Assert.False(pdfOptions.ResourcePolicy.AllowRemoteResourceResolution);
        Assert.False(pdfOptions.ResourcePolicy.AllowLocalFileAccess);
        HtmlRenderRequest pdfRequest = HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, pdfOptions);
        HtmlPdfRenderRequestResult pdfResult = asynchronous
            ? await bundle.RenderToPdfResultAsync(pdfRequest)
            : bundle.RenderToPdfResult(pdfRequest);
        byte[] pdf = pdfResult.ToBytes();
        using var reopened = UglyToad.PdfPig.PdfDocument.Open(pdf);
        Assert.Equal("ARCHIVED", reopened.GetPage(1).Text.Trim());
        Assert.Single(PdfCore.PdfReadDocument.Open(pdf).Pages[0].GetImages());
        Assert.DoesNotContain(pdfResult.RenderResult.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable);
        Assert.Equal(original.ProfileId, pdfResult.RenderResult.Request.ProfileId);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PdfEmbeddedPolicyDeniesArchiveImagesAndDoesNotInvokeAHostResolver(bool asynchronous) {
        using MemoryStream zip = CreateBundle(
            ("index.html", Utf8("<p>TEXT</p><img src='photo.png' width='2' height='2'>")),
            ("photo.png", PdfPngTestImages.CreateRgbPng(2, 2)));
        HtmlSiteBundle bundle = HtmlSiteBundle.Load(zip);
        int hostCalls = 0;
        var options = new HtmlToPdfOptions { ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            ResourceResolver = (request, token) => { hostCalls++; return Task.FromResult<HtmlResolvedResource?>(null); } };
        options.ResourcePolicy.AllowEmbeddedPackageResources = false;
        HtmlRenderRequest request = HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options);
        HtmlPdfRenderRequestResult result = asynchronous
            ? await bundle.RenderToPdfResultAsync(request)
            : bundle.RenderToPdfResult(request);
        Assert.Equal(0, hostCalls);
        Assert.True(result.Output.HasLoss);
        Assert.Empty(PdfCore.PdfReadDocument.Open(result.ToBytes()).Pages[0].GetImages());
        Assert.Contains(result.RenderResult.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable
            || !asynchronous && diagnostic.Code == HtmlRenderDiagnosticCodes.ExternalImagePending);
    }

    [Fact]
    public async Task MissingExternalResourceReportsLossWithoutNetworkOrFilesystemFallback() {
        using MemoryStream zip = CreateBundle(("index.html", Utf8("<p>SAFE</p>"
            + "<img src='https://outside.example.test/missing.png' width='2' height='2'>"
            + "<img src='file:///outside/missing.png' width='2' height='2'>")));
        HtmlSiteBundle bundle = HtmlSiteBundle.Load(zip);
        HtmlRenderRequest request = bundle.CreateRenderRequest(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged));
        HtmlRenderResult result = await HtmlRenderEngine.ExecuteAsync(bundle.HtmlDocument, request);
        Assert.True(result.HasLoss);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable
            && diagnostic.Source == "https://outside.example.test/missing.png");
    }

    [Fact]
    public void RootIndexSelectionIsStableAndOtherAmbiguousBundlesRequireAnEntry() {
        using MemoryStream rooted = CreateBundle(("other.html", Utf8("<p>OTHER</p>")), ("index.html", Utf8("<p>ROOT</p>")));
        Assert.Equal("index.html", HtmlSiteBundle.Load(rooted).EntryPath);
        using MemoryStream ambiguous = CreateBundle(("a.html", Utf8("<p>A</p>")), ("b.html", Utf8("<p>B</p>")));
        Assert.Throws<InvalidDataException>(() => HtmlSiteBundle.Load(ambiguous));
        Assert.Equal("b.html", HtmlSiteBundle.Load(ambiguous, new HtmlSiteBundleOptions { EntryPath = "b.html" }).EntryPath);
        Assert.Throws<InvalidDataException>(() => HtmlSiteBundle.Load(ambiguous, new HtmlSiteBundleOptions { EntryPath = "B.html" }));
    }

    [Theory]
    [InlineData("../outside.css")]
    [InlineData("/absolute.css")]
    [InlineData("C:/absolute.css")]
    [InlineData("styles/../outside.css")]
    public void UnsafeArchivePathsAreRejectedBeforeResourceProjection(string entry) {
        using MemoryStream zip = CreateBundle(("index.html", Utf8("<p>SAFE</p>")), (entry, Utf8("p{}")));
        Assert.Throws<InvalidDataException>(() => HtmlSiteBundle.Load(zip));
    }

    [Fact]
    public void NormalizedAliasesAndSymlinksAreRejected() {
        using MemoryStream aliases = CreateBundle(("index.html", Utf8("<p>SAFE</p>")),
            ("styles/site.css", Utf8("p{}")), ("styles\\site.css", Utf8("p{}")));
        Assert.Throws<InvalidDataException>(() => HtmlSiteBundle.Load(aliases));
        using MemoryStream link = CreateBundle(("index.html", Utf8("<p>SAFE</p>")), ("link.css", Utf8("outside.css")));
        using (var archive = new ZipArchive(link, ZipArchiveMode.Update, leaveOpen: true)) {
            archive.GetEntry("link.css")!.ExternalAttributes = unchecked((int)0xa1ff0000);
        }
        Assert.Throws<InvalidDataException>(() => HtmlSiteBundle.Load(link));
    }

    [Theory]
    [InlineData("./index.html", "./styles/site.css")]
    [InlineData("index.html", "styles\\site.css")]
    public async Task SafeInputPathNormalizationPreservesResourcesInBothLoaders(string entryPath, string cssPath) {
        using MemoryStream zip = CreateBundle(
            (entryPath, Utf8("<link rel='stylesheet' href='styles/site.css'><p>NORMALIZED</p>")),
            (cssPath, Utf8("p{color:#123456}")));
        HtmlSiteBundle synchronous = HtmlSiteBundle.Load(zip);
        HtmlSiteBundle asynchronous = await HtmlSiteBundle.LoadAsync(zip);
        Assert.Equal("index.html", synchronous.EntryPath);
        Assert.Equal(synchronous.EntryNames, asynchronous.EntryNames);
        foreach (HtmlSiteBundle bundle in new[] { synchronous, asynchronous }) {
            HtmlRenderResult rendered = HtmlRenderEngine.Execute(bundle.HtmlDocument, bundle.CreateRenderRequest(
                HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Svg)));
            Assert.Contains("#123456", Encoding.UTF8.GetString(rendered.ExportImage().Bytes), StringComparison.OrdinalIgnoreCase);
        }
    }

    [Theory]
    [InlineData("encoded")]
    [InlineData("entries")]
    [InlineData("entry-bytes")]
    [InlineData("decoded-total")]
    [InlineData("ratio")]
    public async Task InputLimitsRejectZipMetadataAndPayloadExpansion(string limit) {
        byte[] html = Utf8("<p>" + new string('A', 1000) + "</p>");
        byte[] css = Utf8("p{color:red}");
        using MemoryStream zip = CreateBundle(("index.html", html), ("site.css", css));
        var options = new HtmlSiteBundleOptions();
        switch (limit) {
            case "encoded": options.MaximumArchiveBytes = zip.Length - 1; break;
            case "entries": options.MaximumEntryCount = 1; break;
            case "entry-bytes": options.MaximumEntryBytes = html.Length - 1; break;
            case "decoded-total": options.MaximumTotalDecodedBytes = html.Length + css.Length - 1; break;
            case "ratio": options.MaximumCompressionRatio = 1; break;
        }
        await Assert.ThrowsAsync<InvalidDataException>(() => HtmlSiteBundle.LoadAsync(zip, options));
        Assert.True(zip.CanRead);
    }

    [Fact]
    public async Task CallerCancellationAndHtmlInputLimitsArePreserved() {
        using MemoryStream zip = CreateBundle(("index.html", Utf8("<p>INPUT</p>")));
        zip.Position = 7;
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => HtmlSiteBundle.LoadAsync(zip, cancellationToken: cancelled.Token));
        Assert.Equal(7, zip.Position);
        Assert.True(zip.CanRead);
        var htmlOptions = new HtmlConversionDocumentOptions();
        htmlOptions.Limits.MaxInputCharacters = 1;
        await Assert.ThrowsAsync<HtmlDomLimitException>(() => HtmlSiteBundle.LoadAsync(zip, htmlOptions: htmlOptions));
        Assert.Equal(7, zip.Position);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(100)]
    public async Task DeclaredZipLengthCannotTruncateOrOverExpandTheActualEntry(int declaredLength) {
        using MemoryStream valid = CreateBundle(("index.html", Utf8("<p>CONTENTS</p>")));
        byte[] bytes = valid.ToArray();
        // A producer can forge the uncompressed-size field at +24 in the first
        // PK\u0001\u0002 central-directory record without changing its deflate payload.
        int central = -1;
        for (int i = 0; i <= bytes.Length - 28; i++) {
            if (bytes[i] == 0x50 && bytes[i + 1] == 0x4b && bytes[i + 2] == 1 && bytes[i + 3] == 2) {
                central = i;
                break;
            }
        }
        Assert.True(central >= 0);
        for (int i = 0; i < 4; i++) bytes[central + 24 + i] = (byte)(declaredLength >> (8 * i));
        using var malformed = new MemoryStream(bytes);
        await Assert.ThrowsAsync<InvalidDataException>(() => HtmlSiteBundle.LoadAsync(malformed));
        Assert.Throws<InvalidDataException>(() => HtmlSiteBundle.Load(malformed));
    }

    private static byte[] Utf8(string text) => Encoding.UTF8.GetBytes(text);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RasterRenderingUsesTheSameArchiveImageBytes(bool asynchronous) {
        using MemoryStream zip = CreateBundle(
            ("index.html", Utf8("<style>html,body{margin:0}</style><img src='red.png' width='16' height='16'>")),
            ("red.png", PdfPngTestImages.CreateRgbPng(255, 0, 0)));
        HtmlSiteBundle bundle = HtmlSiteBundle.Load(zip);
        HtmlRenderRequest request = bundle.CreateRenderRequest(HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Png,
            new HtmlRenderOptions { Margins = HtmlRenderMargins.All(0) }));
        HtmlRenderResult result = asynchronous
            ? await HtmlRenderEngine.ExecuteAsync(bundle.HtmlDocument, request)
            : HtmlRenderEngine.Execute(bundle.HtmlDocument, request);
        Assert.True(OfficeIMO.Drawing.OfficePngReader.TryDecode(result.ExportImage().Bytes, out var image));
        Assert.Equal(OfficeIMO.Drawing.OfficeColor.FromRgb(255, 0, 0), image!.GetPixel(4, 4));
    }

    [Theory]
    [InlineData(false, "index.html")]
    [InlineData(true, "index.html")]
    [InlineData(false, "unused.css")]
    [InlineData(true, "unused.css")]
    [InlineData(false, "assets/")]
    [InlineData(true, "assets/")]
    public async Task ForgedPrefixSizeAndMatchingChecksumCannotHideDecodedContent(bool stored, string entryName) {
        byte[] prefix = entryName.EndsWith("/", StringComparison.Ordinal) ? Array.Empty<byte>() : Utf8("<p>VISIBLE</p>");
        using var zip = new MemoryStream();
        using (var archive = new ZipArchive(zip, ZipArchiveMode.Create, leaveOpen: true)) {
            if (entryName != "index.html") {
                using Stream root = archive.CreateEntry("index.html").Open();
                byte[] rootBytes = Utf8("<p>ROOT</p>");
                root.Write(rootBytes, 0, rootBytes.Length);
            }
            using Stream payload = archive.CreateEntry(entryName, stored
                ? CompressionLevel.NoCompression : CompressionLevel.Optimal).Open();
            byte[] content = Utf8("<p>VISIBLE</p><p>HIDDEN</p>");
            payload.Write(content, 0, content.Length);
        }
        byte[] bytes = zip.ToArray();
        uint checksum = uint.MaxValue;
        foreach (byte value in prefix) {
            checksum ^= value;
            for (int bit = 0; bit < 8; bit++) checksum = (checksum >> 1) ^ ((checksum & 1) == 0 ? 0 : 0xedb88320U);
        }
        checksum = ~checksum;
        for (int offset = 0; offset <= bytes.Length - 28; offset++) {
            if (bytes[offset] != 0x50 || bytes[offset + 1] != 0x4b) continue;
            bool central = bytes[offset + 2] == 1 && bytes[offset + 3] == 2;
            if (!central) continue;
            int nameLength = bytes[offset + 28] | (bytes[offset + 29] << 8);
            if (Encoding.UTF8.GetString(bytes, offset + 46, nameLength) != entryName) continue;
            int local = BitConverter.ToInt32(bytes, offset + 42);
            for (int i = 0; i < 4; i++) {
                bytes[offset + 16 + i] = bytes[local + 14 + i] = (byte)(checksum >> (8 * i));
                bytes[offset + 24 + i] = bytes[local + 22 + i] = (byte)(prefix.Length >> (8 * i));
            }
        }
        using var malformed = new MemoryStream(bytes);
        Assert.Throws<InvalidDataException>(() => HtmlSiteBundle.Load(malformed));
        await Assert.ThrowsAsync<InvalidDataException>(() => HtmlSiteBundle.LoadAsync(malformed));
    }

    private static MemoryStream CreateBundle(params (string Name, byte[] Bytes)[] entries) {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (var item in entries) {
                ZipArchiveEntry entry = archive.CreateEntry(item.Name, CompressionLevel.Optimal);
                using Stream payload = entry.Open();
                payload.Write(item.Bytes, 0, item.Bytes.Length);
            }
        }
        stream.Position = 0;
        return stream;
    }
}
