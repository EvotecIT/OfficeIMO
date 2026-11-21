using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlPdf_TrimmedFontCachesStayBoundedAndRefillToTheSameOutput() {
        var cache = new OfficeSizedLruCache<int, string>(10L, static (_, _) => 1L);
        for (int key = 0; key < 10; key++) cache.Add(key, "v" + key);
        Assert.True(cache.TryGet(0, out _));

        // Trimming drops the least recently used half and keeps every survivor readable.
        cache.TrimHalf();
        Assert.Equal(5, cache.Count);
        Assert.Equal(5L, cache.RetainedSize);
        Assert.True(cache.TryGet(0, out string kept));
        Assert.Equal("v0", kept);
        Assert.False(cache.TryGet(1, out _));
        for (int key = 10; key < 20; key++) cache.Add(key, "v" + key);
        Assert.Equal(10L, cache.RetainedSize);
        Assert.False(cache.TryGet(0, out _));
        Assert.Equal("v0", kept); // A caller holding an evicted value can still use it.
        for (int key = 10; key < 20; key++) {
            Assert.True(cache.TryGet(key, out string value));
            Assert.Equal("v" + key, value);
        }

        // The process-wide font program and @font-face rule caches refill to identical output.
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A');
        string html = "<style>/*" + new string('x', 70 * 1024) + "*/"
            + "@font-face{font-family:Trimmed;src:url(data:font/ttf;base64," + Convert.ToBase64String(font) + ")}"
            + "p{font-family:Trimmed}</style><p>A</p>";
        byte[] ConvertToPdf() => HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfExportOptions()).ToBytes();
        byte[] cached = ConvertToPdf();
        OfficeMemoryPressure.TrimAll();
        OfficeMemoryPressure.TrimAll();
        Assert.Equal(cached, ConvertToPdf());
        Assert.Equal(cached, ConvertToPdf());
    }
}
