using System.Globalization;
using System.IO.Compression;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfRedactionImageCoverageTests {
    [Fact]
    public void Apply_RewritesRotatedImagePixelsUsingInversePlacementTransform() {
        byte[] source = BuildImagePdf(
            "q\n0 40 -20 0 60 30 cm\n/ImTarget Do\nQ\n",
            "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode",
            Compress(CreateRgbPixels()));
        PdfImagePlacement placement = Assert.Single(PdfImageExtractor.ExtractImagePlacements(source));
        var area = new PdfRedactionArea(1, placement.X, placement.Y, placement.Width / 2D, placement.Height, "rotated-half");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        byte[] pixels = DecodePrimaryImage(redacted, out _);
        Assert.Equal(4, CountBlackPixels(pixels));
        Assert.Single(PdfImageExtractor.ExtractImagePlacements(redacted));
    }

    [Fact]
    public void Apply_EscapesDecodedResourceNamesWhenRewritingImageInvocation() {
        byte[] source = BuildImagePdf(
            "q\n40 0 0 20 20 30 cm\n/Im#20Target Do\nQ\n" +
            "q\n40 0 0 20 100 30 cm\n/Im#20Target Do\nQ\n",
            "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode",
            Compress(CreateRgbPixels()),
            resourceName: "Im#20Target");
        PdfImagePlacement firstPlacement = PdfImageExtractor.ExtractImagePlacements(source)
            .OrderBy(placement => placement.X)
            .First();
        var area = new PdfRedactionArea(1, firstPlacement.X, firstPlacement.Y, firstPlacement.Width / 2D, firstPlacement.Height, "left-half");

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { area });

        string raw = PdfEncoding.Latin1GetString(redacted);
        Assert.Contains("/Im#20TargetRedacted1 Do", raw, StringComparison.Ordinal);
        Assert.Contains("/Im#20Target Do", raw, StringComparison.Ordinal);
        Assert.Equal(2, PdfImageExtractor.ExtractImagePlacements(redacted).Count);
        Assert.Contains(DecodeImages(redacted), pixels => CountBlackPixels(pixels) == 4);
    }

    [Fact]
    public void Apply_EnforcesDecodedImageBudgetBeforeSimplePixelRewrite() {
        byte[] source = BuildImagePdf(
            "q\n40 0 0 20 20 30 cm\n/ImTarget Do\nQ\n",
            "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode",
            Compress(CreateRgbPixels()));

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            PdfRedactionApplier.Apply(
                source,
                new[] { LeftHalfArea(source) },
                new PdfRedactionApplyOptions { MaximumDecodedImageBytes = 8 }));

        Assert.Contains("intersects image placement", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Apply_NormalizesIndexedAndColorKeyImagesBeforePartialRewrite() {
        byte[] indexed = BuildImagePdf(
            "q\n40 0 0 20 20 30 cm\n/ImTarget Do\nQ\n",
            "/ColorSpace [/Indexed /DeviceRGB 1 <FF000000FF00>] /BitsPerComponent 8 /Filter /FlateDecode",
            Compress(new byte[] { 0, 1, 0, 1, 1, 0, 1, 0 }));
        byte[] colorKey = BuildImagePdf(
            "q\n40 0 0 20 20 30 cm\n/ImTarget Do\nQ\n",
            "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Mask [0 0 0 0 255 255]",
            Compress(CreateRgbPixels()));

        byte[] redactedIndexed = RedactLeftHalf(indexed);
        byte[] redactedColorKey = RedactLeftHalf(colorKey);

        Assert.Equal(4, CountBlackPixels(DecodePrimaryImage(redactedIndexed, out _)));
        Assert.Equal(4, CountBlackPixels(DecodePrimaryImage(redactedColorKey, out byte[] colorKeyAlpha)));
        Assert.Equal(8, colorKeyAlpha.Length);
        Assert.All(colorKeyAlpha.Take(2), value => Assert.Equal(255, value));
        Assert.Contains((byte)0, colorKeyAlpha);
    }

    [Fact]
    public void Apply_PreservesExplicitMaskOutsideRewrittenPixels() {
        byte[] source = BuildImagePdf(
            "q\n40 0 0 20 20 30 cm\n/ImTarget Do\nQ\n",
            "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Mask 6 0 R",
            Compress(CreateRgbPixels()),
            "/ImageMask true /BitsPerComponent 1 /Filter /FlateDecode",
            Compress(new byte[] { 0xF0, 0x50 }));

        byte[] redacted = RedactLeftHalf(source);

        byte[] pixels = DecodePrimaryImage(redacted, out byte[] alpha);
        Assert.Equal(4, CountBlackPixels(pixels));
        Assert.Equal(new byte[] { 255, 255, 255, 255, 255, 255, 0, 255 }, alpha);
    }

    [Fact]
    public void Apply_UsesOptionalDecoderForPartialJpegRewrite() {
        byte[] source = BuildImagePdf(
            "q\n40 0 0 20 20 30 cm\n/ImTarget Do\nQ\n",
            "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode",
            new byte[] { 0xFF, 0xD8, 0xFF, 0xD9 });
        var decoder = new TestJpegDecoder();

        byte[] redacted = PdfRedactionApplier.Apply(source, new[] { LeftHalfArea(source) }, new PdfRedactionApplyOptions { ImageDecoder = decoder });

        Assert.True(decoder.WasCalled);
        Assert.Equal(4, CountBlackPixels(DecodePrimaryImage(redacted, out _)));
        Assert.Equal("FlateDecode", Assert.Single(PdfImageExtractor.ExtractImages(redacted)).Filter);
    }

    private static byte[] RedactLeftHalf(byte[] source) => PdfRedactionApplier.Apply(source, new[] { LeftHalfArea(source) });

    private static PdfRedactionArea LeftHalfArea(byte[] source) {
        PdfImagePlacement placement = Assert.Single(PdfImageExtractor.ExtractImagePlacements(source));
        return new PdfRedactionArea(1, placement.X, placement.Y, placement.Width / 2D, placement.Height, "left-half");
    }

    private static byte[] DecodePrimaryImage(byte[] pdf, out byte[] alpha) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf, null);
        PdfExtractedImage extracted = Assert.Single(PdfImageExtractor.ExtractImages(pdf));
        PdfStream image = Assert.IsType<PdfStream>(objects[extracted.ObjectNumber].Value);
        byte[] pixels = StreamDecoder.Decode(image.Dictionary, image.Data, objects);
        alpha = Array.Empty<byte>();
        if (image.Dictionary.Items.TryGetValue("SMask", out PdfObject? maskObject) &&
            maskObject is PdfReference maskReference &&
            objects[maskReference.ObjectNumber].Value is PdfStream maskStream) {
            alpha = StreamDecoder.Decode(maskStream.Dictionary, maskStream.Data, objects);
        }
        return pixels;
    }

    private static byte[][] DecodeImages(byte[] pdf) {
        var (objects, _) = PdfSyntax.ParseObjects(pdf, null);
        return PdfImageExtractor.ExtractImages(pdf)
            .Select(image => Assert.IsType<PdfStream>(objects[image.ObjectNumber].Value))
            .Select(stream => StreamDecoder.Decode(stream.Dictionary, stream.Data, objects))
            .ToArray();
    }

    private static int CountBlackPixels(byte[] rgb) {
        int count = 0;
        for (int offset = 0; offset + 2 < rgb.Length; offset += 3) {
            if (rgb[offset] == 0 && rgb[offset + 1] == 0 && rgb[offset + 2] == 0) count++;
        }
        return count;
    }

    private static byte[] CreateRgbPixels() => new byte[] {
        255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 0,
        255, 0, 255, 0, 255, 255, 128, 64, 32, 240, 240, 240
    };

    private static byte[] BuildImagePdf(string pageContent, string imageEntries, byte[] imageData, string? maskEntries = null, byte[]? maskData = null, string resourceName = "ImTarget") {
        int pageLength = Encoding.ASCII.GetByteCount(pageContent.TrimEnd('\n'));
        using var output = new MemoryStream();
        void Write(string value) { byte[] bytes = Encoding.ASCII.GetBytes(value); output.Write(bytes, 0, bytes.Length); }
        Write(string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 200 120] /Resources << /XObject << /" + resourceName + " 5 0 R >> >> >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + pageLength.ToString(CultureInfo.InvariantCulture) + " >>", "stream", pageContent.TrimEnd('\n'), "endstream", "endobj",
            "5 0 obj", "<< /Type /XObject /Subtype /Image /Width 4 /Height 2 " + imageEntries + " /Length " + imageData.Length.ToString(CultureInfo.InvariantCulture) + " >>", "stream"
        }) + "\n");
        output.Write(imageData, 0, imageData.Length);
        Write("\nendstream\nendobj\n");
        if (maskEntries is not null && maskData is not null) {
            Write("6 0 obj\n<< /Type /XObject /Subtype /Image /Width 4 /Height 2 " + maskEntries + " /Length " + maskData.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
            output.Write(maskData, 0, maskData.Length);
            Write("\nendstream\nendobj\n");
        }
        Write("trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] Compress(byte[] data) {
        using var output = new MemoryStream();
        output.WriteByte(0x78); output.WriteByte(0x9C);
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, true)) deflate.Write(data, 0, data.Length);
        uint a = 1, b = 0;
        for (int i = 0; i < data.Length; i++) { a = (a + data[i]) % 65521; b = (b + a) % 65521; }
        uint adler = (b << 16) | a;
        output.WriteByte((byte)(adler >> 24)); output.WriteByte((byte)(adler >> 16)); output.WriteByte((byte)(adler >> 8)); output.WriteByte((byte)adler);
        return output.ToArray();
    }

    private sealed class TestJpegDecoder : IPdfRedactionImageDecoder {
        public bool WasCalled { get; private set; }

        public bool TryDecode(PdfRedactionImageDecodeRequest request, out PdfRedactionDecodedImage? image) {
            WasCalled = true;
            Assert.Equal("DCTDecode", request.Filter);
            byte[] rgb = CreateRgbPixels();
            byte[] rgba = new byte[request.Width * request.Height * 4];
            for (int pixel = 0; pixel < request.Width * request.Height; pixel++) {
                rgba[pixel * 4] = rgb[pixel * 3];
                rgba[pixel * 4 + 1] = rgb[pixel * 3 + 1];
                rgba[pixel * 4 + 2] = rgb[pixel * 3 + 2];
                rgba[pixel * 4 + 3] = 255;
            }
            image = new PdfRedactionDecodedImage(request.Width, request.Height, rgba);
            return true;
        }
    }
}
