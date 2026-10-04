using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    private const string OpaqueRedFilterMatrix = "0 0 0 0 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1";

    [Fact]
    public void SvgFilterFractionalIntermediateClipDoesNotMoveOpaqueExcessIntoRegion() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><defs>"
            + "<filter id='f' filterUnits='userSpaceOnUse' x='2.25' y='0' width='7.75' height='10'>"
            + "<feColorMatrix/><feOffset dx='2'/></filter></defs>"
            + "<rect width='10' height='10' fill='red' filter='url(#f)'/></svg>";
        var raster = OfficeDrawingRasterRenderer.Render(ReadFilterGraph(svg));
        Assert.InRange(raster.GetPixel(4, 5).A, 1, 254);
        Assert.Equal(255, raster.GetPixel(5, 5).A);
    }

    [Theory]
    [InlineData("")]
    [InlineData("style='mix-blend-mode:multiply'")]
    public void SvgFilterGraphNestedInputKeepsPaintOutsideViewport(string blend) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><defs>"
            + "<filter id='identity'><feColorMatrix/></filter>"
            + "<filter id='move' filterUnits='userSpaceOnUse' x='-5' y='0' width='15' height='10'>"
            + "<feOffset in='SourceGraphic' dx='5'/></filter></defs><g filter='url(#move)'>"
            + "<rect x='-4' y='2' width='3' height='3' fill='lime' filter='url(#identity)' " + blend + "/></g></svg>";
        Assert.True(OfficeDrawingRasterRenderer.Render(ReadFilterGraph(svg)).GetPixel(2, 3).G > 240);
    }

    [Fact]
    public void SvgFilterGraphEmptyChildRetainsOriginalObjectGeometry() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 20'><defs>"
            + "<filter id='empty' filterUnits='userSpaceOnUse' width='0' height='20'><feColorMatrix/></filter>"
            + "<filter id='outer'><feColorMatrix values='" + OpaqueRedFilterMatrix + "'/></filter></defs>"
            + "<g filter='url(#outer)'><rect x='4' y='4' width='4' height='4' filter='url(#empty)'/></g></svg>";
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(ReadFilterGraph(svg));
        Assert.True(raster.GetPixel(5, 5).R > 240);
        Assert.Equal(0, raster.GetPixel(12, 5).A);
    }

    [Fact]
    public void SvgFilterGraphObjectBoundsExcludeMarkerPaint() {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 80 80'><defs>"
            + "<marker id='m' markerUnits='userSpaceOnUse' markerWidth='40' markerHeight='40' refX='0' refY='0' orient='0'>"
            + "<rect width='40' height='40' fill='lime'/></marker>"
            + "<filter id='f'><feColorMatrix values='" + OpaqueRedFilterMatrix + "'/></filter></defs>"
            + "<g filter='url(#f)'><polygon points='10,10 20,10 20,20' marker-end='url(#m)'/></g></svg>";
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(ReadFilterGraph(svg));
        Assert.True(raster.GetPixel(15, 15).R > 240);
        Assert.Equal(0, raster.GetPixel(35, 35).A);
    }

    [Fact]
    public void SvgFilterGraphUndecodableSourceRetainsImageAndReportsUnsupportedFilter() {
        byte[] jp2 = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "red-rgb.jp2"));
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><defs><filter id='f'><feColorMatrix/></filter></defs>"
            + "<g filter='url(#f)'><image width='4' height='4' href='data:image/jp2;base64," + Convert.ToBase64String(jp2) + "'/></g></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(1, unsupported);
        Assert.Equal(jp2, Assert.Single(FilterGraphElements(drawing!).OfType<OfficeDrawingImage>()).Bytes);
    }

    [Theory]
    [InlineData("lime", false)]
    [InlineData("url(#gradient)", false)]
    [InlineData("lime", true)]
    public void SvgFilterGraphChargesPolygonCoverageWorkBeforeRasterization(string fill, bool nestedSurface) {
        string points = string.Join(" ", Enumerable.Range(0, 250).Select(i => {
            double angle = 2D * Math.PI * i / 250D;
            return FormattableString.Invariant($"{100D + 99D * Math.Cos(angle)},{100D + 99D * Math.Sin(angle)}");
        }));
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 200 200'><defs>"
            + "<linearGradient id='gradient'><stop offset='0' stop-color='lime'/><stop offset='1' stop-color='blue'/></linearGradient>"
            + "<filter id='f' filterUnits='userSpaceOnUse' x='0' y='0' width='" + (nestedSurface ? "1" : "200")
            + "' height='" + (nestedSurface ? "1" : "200") + "'><feColorMatrix/></filter></defs>"
            + "<g filter='url(#f)'><polygon points='" + points + "' fill='" + fill + "' "
            + (nestedSurface ? "style='mix-blend-mode:multiply'" : "") + "/></g></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(1, unsupported);
        Assert.Empty(FilterGraphElements(drawing!).OfType<OfficeDrawingImage>());
        Assert.Single(FilterGraphElements(drawing!).OfType<OfficeDrawingShape>());
    }

    [Fact]
    public void SvgFilterGraphChargesNestedEffectAtItsOwnSurfaceSize() {
        string points = string.Join(" ", Enumerable.Range(0, 128).Select(i => {
            double angle = 2D * Math.PI * i / 128D;
            return FormattableString.Invariant($"{512D + 500D * Math.Cos(angle)},{512D + 500D * Math.Sin(angle)}");
        }));
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1024 1024'><defs>"
            + "<filter id='f' filterUnits='userSpaceOnUse' x='0' y='0' width='1' height='1'><feColorMatrix/></filter></defs>"
            + "<g filter='url(#f)'><g style='mix-blend-mode:multiply'><polygon points='" + points
            + "' fill='red'/></g></g></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(1, unsupported);
        Assert.Empty(FilterGraphElements(drawing!).OfType<OfficeDrawingImage>());
    }

    [Fact]
    public void RasterContourRejectsQuadraticSinglePixelSubdivisionWork() {
        var points = Enumerable.Range(0, 6000).Select(i =>
            new OfficePoint(i % 2 == 0 ? 0.1D : 0.9D, 0.1D + 0.8D * i / 6000D)).ToArray();
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(1, 1));

        Assert.Throws<InvalidOperationException>(() => canvas.FillPolygon(points, OfficeColor.Black));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void SvgFilterGraphSourcePolygonCanvasObservesCancellation(int paint) {
        using var cancellation = new CancellationTokenSource();
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(10, 10), null, null, cancellationToken: cancellation.Token);
        cancellation.Cancel();
        var points = new[] {
            new OfficePoint(0, 0), new OfficePoint(9, 0), new OfficePoint(9, 9)
        };
        Assert.Throws<OperationCanceledException>(() => {
            if (paint == 0) canvas.FillPolygon(points, OfficeColor.Black);
            else if (paint == 1) canvas.FillLinearGradientPolygon(points, OfficeLinearGradient.Horizontal(OfficeColor.Black, OfficeColor.White));
            else canvas.FillRadialGradientPolygon(points, OfficeRadialGradient.Centered(OfficeColor.Black, OfficeColor.White));
        });
    }
}
