using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData(8, 100000)]
    [InlineData(12, 100000)]
    [InlineData(16, 2)]
    [InlineData(24, 100000000)]
    [InlineData(28, 100000000)]
    [InlineData(32, 2)]
    public void RenderPage_RejectsUnsafeJpeg2000GeometryBeforeCodec(int fieldOffset, uint value) {
        byte[] payload = ReadScanJpx("rgb");
        int marker = FindJpxCodestream(payload);
        WriteJpxUInt32(payload, marker + fieldOffset, value);
        var codec = new ScanJpxCodec(payload);
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(payload, colorSpace: "/DeviceRGB",
            imageWidth: 1, imageFilterEntry: "/Filter /JPXDecode");
        Assert.False(Assert.Single(PdfPageImageRenderer.RenderPages(pdf,
            options: new PdfPageRenderOptions { ImageCodec = codec })).Succeeded);
        Assert.Equal(0, codec.Calls);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RenderPage_RejectsJpeg2000DimensionsThatDisagreeWithPdf(bool rawCodestream) {
        byte[] payload = ReadScanJpx("rgb");
        if (rawCodestream) payload = payload.Skip(FindJpxCodestream(payload)).ToArray();
        var codec = new ScanJpxCodec(payload);
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(payload, colorSpace: "/DeviceRGB",
            imageWidth: 2, imageFilterEntry: "/Filter /JPXDecode");
        Assert.False(Assert.Single(PdfPageImageRenderer.RenderPages(pdf,
            options: new PdfPageRenderOptions { ImageCodec = codec })).Succeeded);
        Assert.Equal(0, codec.Calls);
    }

    [Fact]
    public void RenderPage_AppliesCallerPixelLimitBeforeJpeg2000Codec() {
        byte[] payload = ReadScanJpx("rgb");
        int marker = FindJpxCodestream(payload);
        int imageHeader = Enumerable.Range(4, payload.Length - 8).Single(index =>
            payload[index] == (byte)'i' && payload[index + 1] == (byte)'h' &&
            payload[index + 2] == (byte)'d' && payload[index + 3] == (byte)'r');
        WriteJpxUInt32(payload, marker + 8, 100000);
        WriteJpxUInt32(payload, marker + 24, 100000);
        WriteJpxUInt32(payload, imageHeader + 8, 100000);
        Assert.True(OfficeJpeg2000Header.TryGetOpaqueDimensions(payload, out _, out int width, out _));
        Assert.Equal(100000, width);
        var codec = new ScanJpxCodec(payload);
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(payload, colorSpace: "/DeviceRGB",
            imageWidth: 100000, imageFilterEntry: "/Filter /JPXDecode");
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf,
            options: new PdfPageRenderOptions { ImageCodec = codec, Dpi = 72, MaxPixelsPerPage = 50000 }));
        Assert.False(result.Succeeded);
        Assert.Equal(0, codec.Calls);
    }

    [Theory]
    [InlineData("0 0 0 20 40 80 cm")]
    [InlineData("20 0 0 0 40 80 cm")]
    [InlineData("20 0 0 20 400 800 cm")]
    [InlineData("0 0 0 0 re W n 20 0 0 20 40 80 cm")]
    public void RenderPage_IgnoresMalformedScansWithoutVisibleArea(string transform) {
        foreach (string filter in new[] { "CCITTFaxDecode", "CCF", "JPXDecode" }) {
            byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(new byte[] { 0 },
                colorSpace: "/DeviceGray", bitsPerComponent: 1, imageWidth: 8,
                imageFilterEntry: "/Filter /" + filter,
                contentStream: "q " + transform + " /Im1 Do Q 1 0 0 rg 10 10 20 20 re f");
            OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
            Assert.Empty(drawing.Images);
            Assert.True(drawing.Elements.Count > 0);
        }
    }

    [Theory]
    [InlineData("CustomJPXDecode")]
    [InlineData("XCCF")]
    [InlineData("CustomCCITTFaxDecode")]
    public void RenderPage_PreservesUnsupportedResourceBehaviorForNonScanFilterNames(string filter) {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(new byte[] { 0 },
            colorSpace: "/DeviceGray", bitsPerComponent: 1, imageWidth: 8,
            imageFilterEntry: "/Filter /" + filter);
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        Assert.True(result.Succeeded);
        Assert.False(Assert.Single(PdfImageExtractor.ExtractImages(pdf)).IsImageFile);
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    public void PackedImageExtractionPropagatesCancellation(bool fax, bool imageMask, bool colorized) {
        var dictionary = new PdfDictionary();
        dictionary.Items["Width"] = new PdfNumber(8);
        dictionary.Items["Height"] = new PdfNumber(1);
        dictionary.Items["BitsPerComponent"] = new PdfNumber(1);
        dictionary.Items["ColorSpace"] = new PdfName("DeviceGray");
        if (imageMask) dictionary.Items["ImageMask"] = new PdfBoolean(true);
        if (fax) dictionary.Items["Filter"] = new PdfName("CCITTFaxDecode");
        var stream = new PdfStream(dictionary, new byte[] { 0 });
        Assert.ThrowsAny<OperationCanceledException>(() => ResourceResolver.BuildExtractedImage(1, "Scan", 1, 0,
            stream, new Dictionary<int, PdfIndirectObject>(), imageMaskColor: OfficeColor.Red,
            colorizeImageMask: colorized, cancellationToken: new CancellationToken(true)));
    }

    private static int FindJpxCodestream(byte[] payload) => Enumerable.Range(0, payload.Length - 3).Single(index =>
        payload[index] == 255 && payload[index + 1] == 79 && payload[index + 2] == 255 && payload[index + 3] == 81);

    private static void WriteJpxUInt32(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)(value >> 24); bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8); bytes[offset + 3] = (byte)value;
    }
}
