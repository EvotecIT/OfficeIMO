using System.Text;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfStampOpacityAndRotationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TextOpacityMultipliesColorAlphaAndPreservesExistingGraphicsStates(bool watermark) {
        byte[] target = PdfStamper.StampText(CanvasPdf(200, 200), "Existing", new PdfTextStampOptions {
            EmbedFont = false, Color = OfficeColor.Blue, Opacity = 0.4,
            X = OfficeLength.Points(20), Y = OfficeLength.Points(140)
        });
        var options = new PdfTextStampOptions {
            EmbedFont = false, Color = OfficeColor.FromRgba(255, 0, 0, 128), Opacity = 0.5,
            X = OfficeLength.Points(20), Y = OfficeLength.Points(80)
        };
        byte[] result = watermark ? PdfStamper.WatermarkText(target, "Stamp alpha", options)
            : PdfStamper.StampText(target, "Stamp alpha", options);

        var text = PdfPageImageRenderer.RenderPage(result).Elements.OfType<OfficeDrawingText>().ToArray();
        Assert.Equal(OfficeColor.FromRgba(0, 0, 255, 102), Assert.Single(text, item => item.Text == "Existing").Color);
        Assert.Equal(OfficeColor.FromRgba(255, 0, 0, 64), Assert.Single(text, item => item.Text == "Stamp alpha").Color);
        Assert.Equal(0.5, options.Opacity);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImageOpacityMultipliesSoftMaskInRenderedPaint(bool watermark) {
        byte[] target = CanvasPdf(100, 100);
        byte[] image = PdfPngTestImages.CreateRgbaPng(255, 0, 0, 128);
        var options = new PdfImageStampOptions {
            X = OfficeLength.Points(20), Y = OfficeLength.Points(20),
            Width = OfficeLength.Points(20), Height = OfficeLength.Points(20), Opacity = 0.5
        };
        byte[] result = watermark ? PdfStamper.WatermarkImage(target, image, options)
            : PdfStamper.StampImage(target, image, options);

        OfficeDrawingImage painted = Assert.Single(PdfPageImageRenderer.RenderPage(result).Images);
        Assert.Equal(0.5, painted.Opacity);
        Assert.Contains("/SMask", PdfEncoding.Latin1GetString(result), StringComparison.Ordinal);
        Assert.True(OfficePngReader.TryDecode(PdfPageImageRenderer.RenderPageAsPng(result), out OfficeRasterImage? raster));
        OfficeColor pixel = raster!.GetPixel(25, 75);
        Assert.Equal(255, pixel.R);
        Assert.InRange(pixel.G, 190, 192);
        Assert.InRange(pixel.B, 190, 192);
    }

    [Fact]
    public void OverlayRotatesPlacementAndCoverClipAroundFrameOriginOnNormalizedPage() {
        byte[] source = CanvasPdf(40, 10, "1 0 0 rg 0 0 40 10 re f");
        byte[] target = CanvasPdf(220, 200, pageEntries: "/CropBox [10 20 210 170] /Rotate 90 /UserUnit 2");
        byte[] result = PdfStamper.OverlayPage(target, source, new PdfPageOverlayOptions {
            X = 60, Y = 80, Width = 40, Height = 20,
            Fit = PdfPageOverlayFit.Cover, RotationDegrees = 90, Opacity = 0.5
        });

        Assert.True(OfficePngReader.TryDecode(PdfPageImageRenderer.RenderPageAsPng(result), out OfficeRasterImage? raster));
        Assert.Equal(300, raster!.Width);
        Assert.Equal(400, raster.Height);
        // The rotated frame occupies x=40..60, y=80..120 in the normalized bottom-left frame.
        OfficeColor inside = raster.GetPixel(50, 290);
        Assert.Equal(255, inside.R);
        Assert.InRange(inside.G, 127, 129);
        Assert.InRange(inside.B, 127, 129);
        Assert.Equal(OfficeColor.White, raster.GetPixel(50, 330)); // rotated Cover clip
        Assert.Equal(OfficeColor.White, raster.GetPixel(80, 310)); // original, unrotated placement
    }

    [Theory]
    [InlineData(-0.01)]
    [InlineData(1.01)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void StampOpacityRejectsValuesOutsideFiniteUnitInterval(double value) {
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfTextStampOptions { Opacity = value });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfImageStampOptions { Opacity = value });
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPageOverlayOptions { RotationDegrees = value });
        }
    }

    private static byte[] CanvasPdf(int width, int height, string content = "", string pageEntries = "") => Encoding.ASCII.GetBytes($"""
        %PDF-1.4
        1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
        2 0 obj << /Type /Pages /Count 1 /Kids [3 0 R] >> endobj
        3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 {width} {height}] {pageEntries} /Resources << >> /Contents 4 0 R >> endobj
        4 0 obj << /Length {content.Length} >> stream
        {content}
        endstream
        endobj
        trailer << /Root 1 0 R >>
        %%EOF
        """);
}
