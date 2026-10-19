using System;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests {

    public sealed class PdfSvgImagePreparationTests {
        [Fact]
        public void SvgImagePreparationRendersPixelsAndEmbedsInPdf() {
            byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20'><rect width='40' height='20' fill='#ff0000'/></svg>");
            Assert.True(PdfDocument.TryPrepareImageBytes(svg, out byte[] png, out OfficeImageInfo? info, out bool transcoded, out string? reason), reason);
            Assert.True(transcoded);
            Assert.Equal(OfficeImageFormat.Png, info!.Format);
            Assert.Equal(40, info.Width);
            Assert.Equal(20, info.Height);
            Assert.True(OfficeRasterImageDecoder.TryDecode(png, out OfficeRasterImage? raster));
            byte[] pixels = raster!.GetPixels();
            Assert.Equal(new byte[] { 255, 0, 0, 255 }, new[] { pixels[0], pixels[1], pixels[2], pixels[3] });
            Assert.Single(PdfImageExtractor.ExtractImages(PdfDocument.Create().Image(svg, 40, 20).ToBytes()));
        }

        [Theory]
        [InlineData("<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20'><foreignObject width='40' height='20'/></svg>")]
        [InlineData("<svg xmlns='http://www.w3.org/2000/svg' width='3000' height='3000'><rect width='3000' height='3000'/></svg>")]
        public void UnsupportedOrOversizedSvgIsRejected(string svg) {
            Assert.False(PdfDocument.TryPrepareImageBytes(Encoding.UTF8.GetBytes(svg), out _, out _, out _, out string? reason));
            Assert.False(string.IsNullOrWhiteSpace(reason));
        }
    }
}
