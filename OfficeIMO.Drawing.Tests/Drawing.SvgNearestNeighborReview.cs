using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeDrawingSvgExporter_DoesNotDiscardSourceSamplesAtScaledOutput() {
        var raster = new OfficeRasterImage(2, 1, OfficeColor.Red);
        raster.SetPixel(1, 0, OfficeColor.Blue);
        byte[] png = OfficePngWriter.Encode(raster);
        var drawing = new OfficeDrawing(1D, 1D);
        drawing.AddImage(
            png,
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 1D, 1D)),
            interpolate: false);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing, 2D, OfficeSvgSizeUnit.Pixel);

        Assert.Contains("fill=\"#FF0000\"", svg, StringComparison.Ordinal);
        Assert.Contains("fill=\"#0000FF\"", svg, StringComparison.Ordinal);
        Assert.Contains("scale(0.5 1)", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_ObservesCancellationDuringNearestNeighborVectorization() {
        var drawing = new OfficeDrawing(20D, 16D);
        drawing.AddImage(
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
        drawing.AddImage(
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
    public void OfficeDrawingSvgExporter_VectorizesOnlyVisibleSourceCrop() {
        var raster = new OfficeRasterImage(1001, 1000, OfficeColor.Black);
        for (int y = 0; y < raster.Height; y++) {
            for (int x = 1; x < raster.Width; x += 2) raster.SetPixel(x, y, OfficeColor.White);
        }
        var drawing = new OfficeDrawing(2D, 1000D);
        drawing.AddImage(
            new byte[] { 1 },
            "image/x-test",
            new OfficeImageProjection(
                new OfficeImagePlacement(0D, 0D, 2D, 1000D),
                new OfficeImageSourceCrop(0.998D, 0D, 0D, 0D)),
            interpolate: false);

        string svg = OfficeDrawingSvgExporter.ToSvg(
            drawing,
            1D,
            OfficeSvgSizeUnit.Pixel,
            new FixedNearestNeighborCodec(raster));

        Assert.Contains("officeimo-image-clip-", svg, StringComparison.Ordinal);
        Assert.True(CountOccurrences(svg, "<rect") < 10000);
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
