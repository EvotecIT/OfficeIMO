using System;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPngPassThroughTests {
    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(4)]
    public void PackedGrayPreservesCompressedRowsAndExtractedPixels(int bitDepth) {
        const int width = 17, height = 5;
        int rowBytes = (width * bitDepth + 7) / 8;
        byte[] rows = new byte[(rowBytes + 1) * height];
        byte[] previous = new byte[rowBytes];
        byte[] expected = new byte[width * height * 4];
        int maxSample = (1 << bitDepth) - 1;
        for (int y = 0; y < height; y++) {
            var current = new byte[rowBytes];
            for (int x = 0; x < width; x++) {
                int sample = (x * 3 + y * 5) & maxSample;
                current[x * bitDepth / 8] |= (byte)(sample << (8 - bitDepth - x * bitDepth % 8));
                int offset = (y * width + x) * 4;
                expected[offset] = expected[offset + 1] = expected[offset + 2] = (byte)(sample * 255 / maxSample);
                expected[offset + 3] = 255;
            }
            int row = y * (rowBytes + 1);
            rows[row] = (byte)y;
            for (int i = 0; i < rowBytes; i++) {
                int left = i == 0 ? 0 : current[i - 1];
                int up = previous[i];
                int upperLeft = i == 0 ? 0 : previous[i - 1];
                int predictor = y switch { 0 => 0, 1 => left, 2 => up, 3 => (left + up) / 2, _ => Paeth(left, up, upperLeft) };
                rows[row + 1 + i] = unchecked((byte)(current[i] - predictor));
            }
            previous = current;
        }
        byte[] png = PdfPngTestImages.CreatePngWithScanlines(width, height, bitDepth, 0, rows);
        Assert.True(PdfWriter.TryGetPngImageData(png, out var prepared, out string? reason), reason);
        Assert.Equal(ReadIdat(png), prepared.Data);
        Assert.Contains("/BitsPerComponent " + bitDepth, prepared.DictionarySuffix);
        Assert.Null(prepared.SoftMask);
        // Stamps and signature appearances use the object-building route.
        PdfStream imageObject = PdfWriter.BuildImageXObject(prepared.CloneForWrite());
        Assert.Equal(bitDepth, imageObject.Dictionary.Get<PdfNumber>("BitsPerComponent")!.Value);
        Assert.Equal(bitDepth, imageObject.Dictionary.Get<PdfDictionary>("DecodeParms")!.Get<PdfNumber>("BitsPerComponent")!.Value);
        byte[] pdf = PdfDocument.Create(doc => doc.Content(content => content.Image(png, 170, 50))).ToBytes();
        var extracted = Assert.Single(PdfDocument.Load(pdf).Images.Extract());
        Assert.Equal(bitDepth, extracted.BitsPerComponent);
        Assert.True(OfficePngReader.TryDecode(extracted.Bytes, out var raster));
        Assert.Equal(expected, raster!.GetPixels());
    }

    [Theory]
    [InlineData(1, 0)]
    [InlineData(2, 0)]
    [InlineData(4, 0)]
    [InlineData(8, 0)]
    [InlineData(8, 2)]
    public void PassThroughRejectsInvalidRowFilter(int bitDepth, int colorType) {
        int colors = colorType == 2 ? 3 : 1;
        byte[] rows = new byte[1 + (17 * bitDepth * colors + 7) / 8];
        rows[0] = 5;
        byte[] png = PdfPngTestImages.CreatePngWithScanlines(17, 1, bitDepth, colorType, rows);
        Assert.False(PdfWriter.TryGetPngImageData(png, out _, out string? reason));
        Assert.Contains("scanline filter", reason);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(4)]
    public void PackedGrayRejectsIncompleteAndExtraDecodedRows(int bitDepth) {
        int expected = 1 + (17 * bitDepth + 7) / 8;
        foreach (int length in new[] { expected - 1, expected + 1 }) {
            byte[] png = PdfPngTestImages.CreatePngWithScanlines(17, 1, bitDepth, 0, new byte[length]);
            Assert.False(PdfWriter.TryGetPngImageData(png, out _, out string? reason));
            Assert.Contains("scanline size", reason);
        }
    }

    private static byte[] ReadIdat(byte[] png) {
        using var bytes = new MemoryStream();
        for (int offset = 8; offset < png.Length;) {
            int length = (png[offset] << 24) | (png[offset + 1] << 16) | (png[offset + 2] << 8) | png[offset + 3];
            if (png[offset + 4] == 'I' && png[offset + 5] == 'D' && png[offset + 6] == 'A' && png[offset + 7] == 'T')
                bytes.Write(png, offset + 8, length);
            offset += length + 12;
        }
        return bytes.ToArray();
    }

    private static int Paeth(int a, int b, int c) {
        int p = a + b - c, da = Math.Abs(p - a), db = Math.Abs(p - b), dc = Math.Abs(p - c);
        return da <= db && da <= dc ? a : db <= dc ? b : c;
    }
}
