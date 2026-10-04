using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfBilevelScanImageTests {
    [Fact]
    public void PreparedBilevelScanPreservesPixelsThroughPngPdfEmbeddingAndExtraction() {
        var source = new OfficeRasterImage(137, 45, OfficeColor.FromRgb(220, 220, 220));
        for (int y = 8; y < 37; y++) for (int x = 13; x < 124; x++) {
            if ((x + y * 3) % 11 < 4) source.SetPixel(x, y, OfficeColor.FromRgb(40, 40, 40));
        }
        byte[] original = source.GetPixels();
        var processed = OfficeScanProcessor.Process(source, new OfficeScanProcessingOptions {
            Deskew = false, NormalizeBackground = false, ColorMode = OfficeScanColorMode.Bilevel
        });
        byte[] png = OfficeRasterImageEncoder.Encode(processed.Image, OfficeImageExportFormat.Png);
        Assert.Equal(1, png[24]);
        Assert.Equal(0, png[25]);
        byte[] pdf = PdfDocument.Create().Image(png, 137, 45).ToBytes();
        var extracted = Assert.Single(PdfImageExtractor.ExtractImages(pdf));
        Assert.Equal("DeviceGray", extracted.ColorSpace);
        Assert.False(extracted.HasTransparencyMask);
        Assert.Equal(processed.Image.Width, extracted.Width);
        Assert.Equal(processed.Image.Height, extracted.Height);
        Assert.True(OfficeRasterImageDecoder.TryDecode(extracted.Bytes, out var decoded));
        Assert.Equal(processed.Image.GetPixels(), decoded!.GetPixels());
        Assert.Equal(original, source.GetPixels());
    }
}
