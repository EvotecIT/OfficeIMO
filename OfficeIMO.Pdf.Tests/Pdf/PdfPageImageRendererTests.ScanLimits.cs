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
    [InlineData("0 0 0 0 re W n 0 20 -20 0 60 80 cm")]
    [InlineData("500 500 20 20 re W n 0 20 -20 0 60 80 cm")]
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
    [InlineData("Custom,JPXDecode")]
    [InlineData("JPXDecode#20")]
    [InlineData("CCF,Custom")]
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

    [Fact]
    public void Jpeg2000ExtractionPropagatesCancellation() {
        var dictionary = new PdfDictionary();
        dictionary.Items["Width"] = new PdfNumber(1);
        dictionary.Items["Height"] = new PdfNumber(1);
        dictionary.Items["ColorSpace"] = new PdfName("DeviceRGB");
        dictionary.Items["Filter"] = new PdfName("JPXDecode");
        var stream = new PdfStream(dictionary, ReadScanJpx("rgb"));
        Assert.ThrowsAny<OperationCanceledException>(() => ResourceResolver.BuildExtractedImage(1, "Scan", 1, 0,
            stream, new Dictionary<int, PdfIndirectObject>(), cancellationToken: new CancellationToken(true)));
    }

    [Fact]
    public void RenderPage_NormalizesIndexedCcittAndDecodesWrappedJpeg2000() {
        byte[] fax = BuildSingleStreamPdfWithBinaryImageXObject(PdfFaxDecodeTests.Pack("001 00110101 000101"),
            colorSpace: "[/Indexed /DeviceRGB 1 <000000FFFFFF>]", bitsPerComponent: 1, imageWidth: 8,
            imageFilterEntry: "/Filter /CCITTFaxDecode /DecodeParms << /K -1 /Columns 8 /Rows 1 /EndOfBlock false >>");
        OfficeDrawingImage image = Assert.Single(PdfPageImageRenderer.RenderPage(fax).Images);
        Assert.True(OfficeRasterImageDecoder.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.Equal(OfficeColor.Black, raster!.GetPixel(0, 0));
        byte[] jpx = ReadScanJpx("rgb");
        var codec = new ScanJpxCodec(jpx);
        byte[] wrapped = BuildSingleStreamPdfWithBinaryImageXObject(CompressWithDeflate(jpx),
            colorSpace: "/DeviceRGB", imageWidth: 1, imageFilterEntry: "/Filter [/FlateDecode /JPXDecode]");
        Assert.True(Assert.Single(PdfPageImageRenderer.RenderPages(wrapped,
            options: new PdfPageRenderOptions { ImageCodec = codec })).Succeeded);
        Assert.True(codec.Calls > 0);
    }

    [Fact]
    public void IndexedPaletteNormalizationObservesCancellation() {
        var palette = new PdfArray();
        palette.Items.Add(new PdfName("Indexed")); palette.Items.Add(new PdfName("DeviceRGB"));
        palette.Items.Add(new PdfNumber(1)); palette.Items.Add(new PdfStringObj(new byte[] { 0, 0, 0, 255, 255, 255 }));
        var stream = new PdfStream(new PdfDictionary(), new byte[] { 0 });
        Assert.ThrowsAny<OperationCanceledException>(() => PdfIndexedImageNormalizer.TryBuildPngFile(
            palette, 8, 1, 1, stream, new Dictionary<int, PdfIndirectObject>(), 1024,
            out _, new CancellationToken(true)));
    }

    private static int FindJpxCodestream(byte[] payload) => Enumerable.Range(0, payload.Length - 3).Single(index =>
        payload[index] == 255 && payload[index + 1] == 79 && payload[index + 2] == 255 && payload[index + 3] == 81);

    private static void WriteJpxUInt32(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)(value >> 24); bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8); bytes[offset + 3] = (byte)value;
    }
}
