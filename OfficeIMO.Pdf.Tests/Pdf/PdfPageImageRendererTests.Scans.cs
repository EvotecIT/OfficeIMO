using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData(1, 0x40, "", 0, 255)]
    [InlineData(1, 0x40, " /Decode [1 0]", 255, 0)]
    [InlineData(2, 0x30, "", 0, 255)]
    [InlineData(4, 0x0F, "", 0, 255)]
    public void RenderPage_ProjectsPackedGrayThroughDecodePipeline(int bits, int packed, string extra, byte first, byte second) {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(new[] { (byte)packed },
            colorSpace: "/DeviceGray", bitsPerComponent: bits, imageWidth: 2,
            extraImageEntries: extra, imageFilterEntry: "");
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeDrawingImage image = Assert.Single(drawing.Images);
        Assert.True(OfficeRasterImageDecoder.TryDecode(image.Bytes, out OfficeRasterImage? decoded));
        Assert.Equal(OfficeColor.FromRgb(first, first, first), decoded!.GetPixel(0, 0));
        Assert.Equal(OfficeColor.FromRgb(second, second, second), decoded.GetPixel(1, 0));
        Assert.DoesNotContain(PdfPageImageRenderer.RenderPages(pdf).Single().CapabilityDiagnostics,
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Fact]
    public void RenderPage_PreservesPackedGrayColorKeyMask() {
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(new byte[] { 0x40 },
            colorSpace: "/DeviceGray", bitsPerComponent: 1, imageWidth: 2,
            extraImageEntries: " /Mask [1 1]", imageFilterEntry: "");
        OfficeDrawingImage image = Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Images);
        Assert.True(OfficeRasterImageDecoder.TryDecode(image.Bytes, out OfficeRasterImage? decoded));
        Assert.Equal(OfficeColor.Black, decoded!.GetPixel(0, 0));
        Assert.Equal(0, decoded.GetPixel(1, 0).A);
    }

    [Fact]
    public async Task Ocr_RendersCcittAndRejectsMalformedScanBeforeProviderCall() {
        byte[] valid = BuildSingleStreamPdfWithBinaryImageXObject(PdfFaxDecodeTests.Pack("001 00110101 000101"),
            colorSpace: "/DeviceGray", bitsPerComponent: 1, imageWidth: 8,
            imageFilterEntry: "/Filter /CCITTFaxDecode /DecodeParms << /K -1 /Columns 8 /Rows 1 /EndOfBlock false >>");
        int calls = 0;
        var engine = new DelegateOcrEngine("scan-proof", (request, token) => {
            calls++;
            Assert.True(OfficeRasterImageDecoder.TryDecode(request.Payload, out OfficeRasterImage? raster));
            Assert.Equal(OfficeColor.Black, raster!.GetPixel(50, 110));
            return Task.FromResult(new OcrResult());
        });
        await PdfDocument.Load(valid).ReadWithOcrAsync(engine, new PdfOcrMergeOptions { Dpi = 72 });
        Assert.Equal(1, calls);
        byte[] invalid = BuildSingleStreamPdfWithBinaryImageXObject(new byte[] { 0 },
            colorSpace: "/DeviceGray", bitsPerComponent: 1, imageWidth: 8,
            imageFilterEntry: "/Filter /CCITTFaxDecode /DecodeParms << /K -1 /Columns 8 /Rows 1 /EndOfBlock false >>");
        PdfPageRenderResult failed = Assert.Single(PdfPageImageRenderer.RenderPages(invalid));
        Assert.False(failed.Succeeded);
        await Assert.ThrowsAsync<NotSupportedException>(() => PdfDocument.Load(invalid).ReadWithOcrAsync(engine));
        Assert.Equal(1, calls);
    }

    [Fact]
    public async Task Ocr_UsesJpeg2000CodecAndDoesNotSendBlankPageWhenCodecIsMissing() {
        byte[] payload = ReadScanJpx("rgb");
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(payload,
            colorSpace: "/DeviceRGB", imageWidth: 1, imageFilterEntry: "/Filter /JPXDecode");
        var codec = new ScanJpxCodec(payload);
        int calls = 0;
        var engine = new DelegateOcrEngine("codec-proof", (request, token) => {
            calls++;
            Assert.True(OfficeRasterImageDecoder.TryDecode(request.Payload, out OfficeRasterImage? raster));
            Assert.Equal(OfficeColor.Red, raster!.GetPixel(50, 110));
            return Task.FromResult(new OcrResult());
        });
        await Assert.ThrowsAsync<NotSupportedException>(() => PdfDocument.Load(pdf).ReadWithOcrAsync(engine));
        Assert.Equal(0, calls);
        PdfOcrMergeResult result = await PdfDocument.Load(pdf).ReadWithOcrAsync(engine,
            new PdfOcrMergeOptions { Dpi = 72, ImageCodec = codec });
        Assert.Equal(1, calls);
        Assert.True(codec.Calls > 0);
        Assert.Contains(result.Pages.Single().Diagnostics, diagnostic => diagnostic.StartsWith(PdfRenderCapabilities.OptionalImageCodecId));
    }

    [Theory]
    [InlineData("rgba", "")]
    [InlineData("rgba", " /SMaskInData 0")]
    [InlineData("rgba", " /SMaskInData 1")]
    [InlineData("rgba", " /SMaskInData 2")]
    [InlineData("rgb", " /SMaskInData 1")]
    [InlineData("rgb", " /SMaskInData 2")]
    public async Task Ocr_RejectsJpeg2000OpacityBeforeRenderingOrCallingProvider(string mode, string mask) {
        byte[] payload = ReadScanJpx(mode);
        byte[] pdf = BuildSingleStreamPdfWithBinaryImageXObject(payload,
            colorSpace: "/DeviceRGB", imageWidth: 1, imageFilterEntry: "/Filter /JPXDecode",
            extraImageEntries: mask);
        var codec = new ScanJpxCodec(payload);
        Assert.False(Assert.Single(PdfPageImageRenderer.RenderPages(pdf,
            options: new PdfPageRenderOptions { ImageCodec = codec })).Succeeded);
        int calls = 0;
        var engine = new DelegateOcrEngine("alpha-proof", (request, token) => {
            calls++;
            return Task.FromResult(new OcrResult());
        });
        await Assert.ThrowsAsync<NotSupportedException>(() => PdfDocument.Load(pdf).ReadWithOcrAsync(engine,
            new PdfOcrMergeOptions { ImageCodec = codec }));
        Assert.Equal(0, calls);
        Assert.Equal(0, codec.Calls);
    }

    [Fact]
    public void Jpeg2000Header_RejectsTruncatedAndOversizedBoxesAndRecognizesOpaqueCodestream() {
        byte[] payload = ReadScanJpx("rgb");
        Assert.True(OfficeJpeg2000Header.TryGetOpaqueComponents(payload, out int components));
        Assert.Equal(3, components);
        for (int length = 0; length < payload.Length; length++) {
            Assert.False(OfficeJpeg2000Header.TryGetOpaqueComponents(payload.Take(length).ToArray(), out _));
        }
        byte[] oversized = (byte[])payload.Clone();
        for (int i = 12; i < 16; i++) oversized[i] = 255;
        Assert.False(OfficeJpeg2000Header.TryGetOpaqueComponents(oversized, out _));
        int marker = Enumerable.Range(0, payload.Length - 3).Single(index =>
            payload[index] == 255 && payload[index + 1] == 79 && payload[index + 2] == 255 && payload[index + 3] == 81);
        Assert.True(OfficeJpeg2000Header.TryGetOpaqueComponents(payload.Skip(marker).ToArray(), out components));
        Assert.Equal(3, components);
        Assert.False(OfficeJpeg2000Header.TryGetOpaqueComponents(ReadScanJpx("rgba"), out _));
    }

    private static byte[] ReadScanJpx(string mode) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,
        "Pdf", "Fixtures", "Interoperability", "Scans", "red-" + mode + ".jp2"));

    private sealed class ScanJpxCodec : IOfficeRasterImageCodec {
        private readonly byte[] _expected;
        internal ScanJpxCodec(byte[] expected) { _expected = expected; }
        internal int Calls { get; private set; }
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            Assert.Equal("image/jp2", contentType);
            Assert.Equal(_expected, encodedBytes);
            Calls++;
            image = new OfficeRasterImage(1, 1, OfficeColor.Red);
            return true;
        }
    }
}
