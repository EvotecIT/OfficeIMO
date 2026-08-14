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
}
