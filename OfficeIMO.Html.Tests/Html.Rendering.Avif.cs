using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("avif-opaque", null)]
    [InlineData("avif-opaque", "image/avif")]
    [InlineData("avif-opaque", " IMAGE/AVIF ; codecs=av01 ")]
    [InlineData("avif-alpha", null)]
    [InlineData("avif-alpha", "image/avif")]
    [InlineData("avif-alpha", " IMAGE/AVIF ; codecs=av01 ")]
    [InlineData("avif-monochrome-full", "image/avif")]
    [InlineData("avif-monochrome-limited", "image/avif")]
    [InlineData("avif-monochrome-full-alpha", "image/avif")]
    [InlineData("avif-monochrome-limited-alpha", "image/avif")]
    [InlineData("avif-main10-420-full", "image/avif")]
    [InlineData("avif-main10-420-limited", "image/avif")]
    [InlineData("avif-main10-420-full-alpha", "image/avif")]
    [InlineData("avif-main10-420-limited-alpha", "image/avif")]
    [InlineData("avif-main10-mono-full", "image/avif")]
    [InlineData("avif-main10-mono-limited", "image/avif")]
    [InlineData("avif-main10-mono-full-alpha", "image/avif")]
    [InlineData("avif-main10-mono-limited-alpha", "image/avif")]
    [InlineData("avif-main8-fullheader-420-full-alpha", "image/avif")]
    [InlineData("avif-main10-fullheader-mono-limited-alpha", "image/avif")]
    public void HtmlRender_AvifPreservesIndependentPixelsInScreenAndPdf(string name, string? pictureType) {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "Html", "Qualification",
            name.Contains("-fullheader-") ? "AvifFullHeaders" : name.StartsWith("avif-main10-", StringComparison.Ordinal) ? "AvifMain10" : name.StartsWith("avif-monochrome-", StringComparison.Ordinal) ? "AvifMonochrome" : "StaticPdfGaps");
        byte[] bytes = File.ReadAllBytes(Path.Combine(root, name + ".avif"));
        byte[] reference = File.ReadAllBytes(Path.Combine(root, name + ".rgba"));
        string source = "data:image/avif;base64," + Convert.ToBase64String(bytes);
        string html = pictureType == null ? "<img width='49' height='33' src='" + source + "'>"
            : "<picture><source type='" + pictureType + "' srcset='" + source + " 1x' width='49' height='33'>"
              + "<img width='4' height='2' src='data:image/png;base64,"
              + Convert.ToBase64String(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(4, 2)) + "'></picture>";
        var document = HtmlConversionDocument.Parse(html);
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous, ViewportWidth = 49, ViewportHeight = 33,
            Margins = HtmlRenderMargins.All(0), Scale = 1,
            ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile()
        };
        var rendered = HtmlRenderTestDriver.Render(document, options);
        Assert.Empty(rendered.Diagnostics);
        var visual = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderImage>());
        Assert.Equal(49, visual.Width);
        Assert.Equal(33, visual.Height);
        Assert.True(OfficeRasterImageDecoder.TryDecode(visual.Bytes, out var raster));
        AssertAvifPixels(reference, raster!,name.StartsWith("avif-main10-",StringComparison.Ordinal)?1:3);
        var result = document.RenderToPdfResult(HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged,
            HtmlRenderEncoder.Pdf, new HtmlToPdfOptions(options)));
        Assert.Empty(result.Output.Warnings);
        var embedded = Assert.Single(OfficeIMO.Pdf.PdfImageExtractor.ExtractImages(result.ToBytes()));
        Assert.True(OfficeRasterImageDecoder.TryDecode(embedded.Bytes, out var extracted));
        AssertAvifPixels(reference, extracted!,name.StartsWith("avif-main10-",StringComparison.Ordinal)?1:3);
    }

    [Theory]
    [InlineData(HtmlCssMediaContext.Print, true)]
    [InlineData(HtmlCssMediaContext.Screen, false)]
    public async Task HtmlRenderAsync_AvifPictureUsesOnlyItsActiveResourceAndDimensionHints(HtmlCssMediaContext media, bool avifSelected) {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "Html", "Qualification", "StaticPdfGaps");
        byte[] avif = File.ReadAllBytes(Path.Combine(root, "avif-alpha.avif"));
        byte[] png = OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(4, 2);
        var requested = new List<Uri>();
        var document = HtmlConversionDocument.Parse("<picture><source media='print' type='image/avif' "
            + "srcset='https://assets.example.test/selected.avif 1x' width='49' height='33'>"
            + "<img src='https://assets.example.test/fallback.png' width='4' height='2'></picture>",
            new HtmlConversionDocumentOptions {
                Profile = avifSelected ? HtmlConversionProfile.HighFidelityPrint : HtmlConversionProfile.Semantic,
                UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
            });
        string selectedSource = "https://assets.example.test/" + (avifSelected ? "selected.avif" : "fallback.png");
        string inactiveSource = "https://assets.example.test/" + (avifSelected ? "fallback.png" : "selected.avif");
        Assert.Contains(document.ResourceManifest.Resources, resource => resource.Source == selectedSource);
        Assert.DoesNotContain(document.ResourceManifest.Resources, resource => resource.Source == inactiveSource);
        var options = new HtmlRenderOptions {
            Mode = avifSelected ? HtmlRenderMode.Paged : HtmlRenderMode.Continuous,
            ViewportWidth = 100, ViewportHeight = 100, Margins = HtmlRenderMargins.All(0),
            ResourceUrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
            ResourceResolver = (request, _) => {
                requested.Add(request.Uri);
                bool selected = request.Uri.AbsolutePath.EndsWith(".avif", StringComparison.Ordinal);
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(selected ? avif : png,
                    selected ? "image/avif" : "image/png"));
            }
        };
        Assert.Equal(media, options.MediaContext);
        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(document, options);
        Assert.Empty(rendered.Diagnostics);
        Assert.Equal(new[] { new Uri(selectedSource) }, requested);
        var image = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderImage>());
        Assert.Equal(avifSelected ? 49 : 4, image.Width);
        Assert.Equal(avifSelected ? 33 : 2, image.Height);
        Assert.True(OfficeRasterImageDecoder.TryDecode(image.Bytes, out var decoded));
        if (avifSelected) AssertAvifPixels(File.ReadAllBytes(Path.Combine(root, "avif-alpha.rgba")), decoded!);
    }

    private static void AssertAvifPixels(byte[] reference, OfficeRasterImage image,int rgbTolerance=3) {
        Assert.Equal(49, image.Width); Assert.Equal(33, image.Height);
        Assert.Equal(reference.Length, image.PixelBuffer.Length);
        for (int i = 0; i < reference.Length; i++)
            Assert.InRange(Math.Abs(reference[i] - image.PixelBuffer[i]), 0, i % 4 == 3 ? 0 : rgbTolerance);
    }

    [Theory]
    [InlineData("avif-opaque")]
    [InlineData("avif-alpha")]
    public void HtmlToWord_AvifPictureStoresOfficeCompatiblePngWithIndependentPixels(string name) {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "Html", "Qualification", "StaticPdfGaps");
        string source = "data:image/avif;base64," + Convert.ToBase64String(File.ReadAllBytes(Path.Combine(root, name + ".avif")));
        var document = HtmlConversionDocument.Parse("<picture><source type='image/avif' srcset='" + source
            + "'><img width='49' height='33' alt='photo' src='data:image/png;base64,"
            + Convert.ToBase64String(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(4, 2)) + "'></picture>");
        var result = document.ToWordDocumentResult();
        using var word = result.Value;
        using var stream = new MemoryStream(word.ToBytes());
        using var package = WordprocessingDocument.Open(stream, false);
        ImagePart part = Assert.Single(package.MainDocumentPart!.ImageParts);
        Assert.Equal("image/png", part.ContentType);
        using var pixels = new MemoryStream();
        using (Stream content = part.GetStream()) content.CopyTo(pixels);
        Assert.True(OfficeRasterImageDecoder.TryDecode(pixels.ToArray(), out var decoded));
        AssertAvifPixels(File.ReadAllBytes(Path.Combine(root, name + ".rgba")), decoded!);
    }
}
