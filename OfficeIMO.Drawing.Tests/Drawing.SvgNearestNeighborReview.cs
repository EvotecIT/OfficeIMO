using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeDrawingSvgExporter_DoesNotDiscardSourceSamplesAtScaledOutput() {
        var raster = new OfficeRasterImage(2, 1, OfficeColor.Red);
        raster.SetPixel(1, 0, OfficeColor.FromRgba(0, 0, 255, 127));
        byte[] png = OfficePngWriter.Encode(raster);
        var drawing = new OfficeDrawing(1D, 1D);
        drawing.AddImageWithInterpolation(
            png,
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 1D, 1D)),
            interpolate: false);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing, 2D, OfficeSvgSizeUnit.Pixel);

        XElement image = Assert.Single(XElement.Parse(svg).Descendants().Where(e => e.Name.LocalName == "image"));
        Assert.Equal("optimizeSpeed", image.Attribute("image-rendering")?.Value);
        Assert.Equal("image-rendering:pixelated", image.Attribute("style")?.Value);
        byte[] embedded = Convert.FromBase64String(image.Attribute("href")!.Value.Split(',')[1]);
        Assert.Equal(png, embedded);
        Assert.True(OfficePngReader.TryDecode(embedded, out OfficeRasterImage? decoded));
        Assert.Equal(raster.GetPixels(), decoded!.GetPixels());
    }

    [Fact]
    public void OfficeDrawingSvgExporter_ObservesCancellationDuringNearestNeighborTranscoding() {
        var drawing = new OfficeDrawing(20D, 16D);
        drawing.AddImageWithInterpolation(
            new byte[] { 1, 2, 3 },
            "image/x-test",
            new OfficeImageProjection(new OfficeImagePlacement(4D, 3D, 8D, 6D)),
            interpolate: false);
        using var cancellation = new CancellationTokenSource();

        Assert.Throws<OperationCanceledException>(() => OfficeDrawingSvgExporter.ToSvgBytes(
            drawing,
            1D,
            OfficeSvgSizeUnit.Pixel,
            new CancelingNearestNeighborCodec(cancellation),
            resourceIdPrefix: null,
            cancellationToken: cancellation.Token));
    }

    [Fact]
    public void OfficeDrawingSvgExporter_SkipsTransparentNearestNeighborImageBeforeDecoding() {
        var drawing = new OfficeDrawing(1D, 1D);
        drawing.AddImageWithInterpolation(
            new byte[] { 1, 2, 3 },
            "image/x-test",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 1D, 1D)),
            interpolate: false,
            opacity: 0D);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing, 1D, OfficeSvgSizeUnit.Pixel);

        Assert.DoesNotContain("<image", svg, StringComparison.Ordinal);
        Assert.DoesNotContain("<rect", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_RetainsSourceCropWhenEmbeddingTranscodedNearestNeighborImage() {
        var raster = new OfficeRasterImage(4, 2, OfficeColor.Black);
        for (int y = 0; y < raster.Height; y++) {
            for (int x = 1; x < raster.Width; x += 2) raster.SetPixel(x, y, OfficeColor.White);
        }
        var drawing = new OfficeDrawing(2D, 2D);
        drawing.AddImageWithInterpolation(
            new byte[] { 1 },
            "image/x-test",
            new OfficeImageProjection(
                new OfficeImagePlacement(0D, 0D, 2D, 2D),
                new OfficeImageSourceCrop(0.5D, 0D, 0D, 0D)),
            interpolate: false);

        string svg = OfficeDrawingSvgExporter.ToSvg(
            drawing,
            1D,
            OfficeSvgSizeUnit.Pixel,
            new FixedNearestNeighborCodec(raster));

        var root = XElement.Parse(svg);
        XElement image = Assert.Single(root.Descendants().Where(e => e.Name.LocalName == "image"));
        Assert.Equal("-2", image.Attribute("x")?.Value);
        Assert.Equal("4", image.Attribute("width")?.Value);
        Assert.NotNull(image.Attribute("clip-path"));
        XElement clip = Assert.Single(root.Descendants().Where(e => e.Name.LocalName == "clipPath"));
        Assert.Equal("2", Assert.Single(clip.Elements()).Attribute("width")?.Value);
        Assert.True(OfficePngReader.TryDecode(Convert.FromBase64String(image.Attribute("href")!.Value.Split(',')[1]), out OfficeRasterImage? embedded));
        Assert.Equal(raster.GetPixels(), embedded!.GetPixels());
    }

    private sealed class CancelingNearestNeighborCodec : IOfficeRasterImageCodec {
        private readonly CancellationTokenSource _cancellation;

        internal CancelingNearestNeighborCodec(CancellationTokenSource cancellation) {
            _cancellation = cancellation;
        }

        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            _cancellation.Cancel();
            image = new OfficeRasterImage(64, 64, OfficeColor.White);
            return true;
        }
    }

    private sealed class FixedNearestNeighborCodec : IOfficeRasterImageCodec {
        private readonly OfficeRasterImage _image;

        internal FixedNearestNeighborCodec(OfficeRasterImage image) {
            _image = image;
        }

        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            image = _image;
            return true;
        }
    }
}
