using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void SvgFilterGraphRendersNamedSourceAlphaShadowAndColorMatrixBlend() {
        const string svg = """
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 360 180">
              <defs>
                <filter id="shadow" x="-30%" y="-30%" width="180%" height="180%">
                  <feGaussianBlur in="SourceAlpha" stdDeviation="4" result="blur"/>
                  <feOffset in="blur" dx="12" dy="8" result="off"/>
                  <feComposite in="SourceGraphic" in2="off" operator="over"/>
                </filter>
                <filter id="color">
                  <feColorMatrix type="matrix" values="0 0 1 0 0 0 1 0 0 0 1 0 0 0 0 0 0 0 1 0" result="matrix"/>
                  <feBlend in="SourceGraphic" in2="matrix" mode="multiply"/>
                </filter>
              </defs>
              <rect x="35" y="35" width="100" height="90" fill="#e03020" filter="url(#shadow)"/>
              <rect x="210" y="35" width="100" height="90" fill="#c040a0" filter="url(#color)"/>
            </svg>
            """;
        OfficeDrawing drawing = ReadFilterGraph(svg);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Equal(OfficeColor.FromRgb(224, 48, 32), raster.GetPixel(60, 60));
        OfficeColor shadow = raster.GetPixel(138, 80);
        Assert.InRange(shadow.A, (byte)180, (byte)255);
        Assert.Equal(0, shadow.R); Assert.Equal(0, shadow.G); Assert.Equal(0, shadow.B);
        OfficeColor faded = raster.GetPixel(152, 80);
        Assert.InRange(faded.A, (byte)1, (byte)80);
        // Independent linear-light multiplication of sRGB #C040A0 and #A040C0.
        OfficeColor mixed = raster.GetPixel(250, 70);
        Assert.InRange(mixed.R, (byte)118, (byte)120);
        Assert.InRange(mixed.G, (byte)8, (byte)10);
        Assert.InRange(mixed.B, (byte)118, (byte)120);
        Assert.Equal(255, mixed.A);
    }

    [Theory]
    [InlineData("", 188)]
    [InlineData("color-interpolation-filters='sRGB'", 128)]
    public void SvgFilterGraphUsesDeclaredColorSpaceAndUnpremultipliedMatrix(string space, int red) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><defs><filter id='f' " + space + ">"
            + "<feColorMatrix values='0.5 0 0 0 0 0 1 0 0 0 0 0 1 0 0 0 0 0 0.5 0'/></filter></defs>"
            + "<rect x='2' y='2' width='6' height='6' fill='red' fill-opacity='0.5' filter='url(#f)'/></svg>";
        OfficeColor pixel = OfficeDrawingRasterRenderer.Render(ReadFilterGraph(svg)).GetPixel(4, 4);
        Assert.InRange((int)pixel.R, red - 1, red + 1);
        Assert.InRange(pixel.A, (byte)63, (byte)65);
        Assert.Equal(0, pixel.G); Assert.Equal(0, pixel.B);
    }

    [Fact]
    public void SvgFilterGraphClipsIntermediateResultsBeforeTheyAreMovedBack() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 6'><defs>"
            + "<filter id='f' filterUnits='userSpaceOnUse' x='0' y='0' width='3' height='6'>"
            + "<feOffset in='SourceGraphic' dx='2' result='moved'/><feOffset in='moved' dx='-2'/></filter></defs>"
            + "<rect width='4' height='4' fill='red' filter='url(#f)'/></svg>";
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(ReadFilterGraph(svg));
        Assert.True(raster.GetPixel(0, 1).R > 240);
        Assert.Equal(0, raster.GetPixel(2, 1).A);
    }

    [Fact]
    public void SvgFilterGraphRetainsInputOutsideViewportUntilTheFinalViewportClip() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><defs>"
            + "<filter id='f' filterUnits='userSpaceOnUse' x='-5' y='0' width='15' height='10'>"
            + "<feOffset in='SourceGraphic' dx='5'/></filter></defs>"
            + "<rect x='-4' y='2' width='3' height='3' fill='lime' filter='url(#f)'/></svg>";
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(ReadFilterGraph(svg));
        Assert.True(raster.GetPixel(2, 3).G > 240);
        Assert.Equal(0, raster.GetPixel(6, 3).A);
    }

    [Fact]
    public void SvgFilterGraphPreservesFractionalRegionClipWhenScaled() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 6'><defs>"
            + "<filter id='f' filterUnits='userSpaceOnUse' x='2.25' y='1' width='2.5' height='4'>"
            + "<feColorMatrix/></filter></defs><rect x='2' y='1' width='6' height='4' fill='red' filter='url(#f)'/></svg>";
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(ReadFilterGraph(svg), 4D);
        Assert.True(raster.GetPixel(12, 8).R > 240);
        Assert.Equal(0, raster.GetPixel(8, 8).A);
        Assert.Equal(0, raster.GetPixel(19, 8).A);
    }

    [Fact]
    public void SvgFilterGraphSupportsNestedFilteredGeometryAndTinyDeviation() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 20'><defs>"
            + "<filter id='identity'><feColorMatrix/></filter>"
            + "<filter id='blur'><feGaussianBlur in='SourceGraphic' stdDeviation='1e-200'/></filter></defs>"
            + "<g filter='url(#identity)'><rect x='4' y='4' width='4' height='4' fill='lime' filter='url(#blur)'/></g></svg>";
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(ReadFilterGraph(svg));
        Assert.True(raster.GetPixel(5, 5).G > 240);
        Assert.Equal(0, raster.GetPixel(2, 5).A);
    }

    [Fact]
    public void SvgFilterGraphUsesViewBoxOriginAndScaledOffsetCoordinates() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='10 0 40 20'><defs>"
            + "<filter id='f' filterUnits='userSpaceOnUse' x='10' y='0' width='20' height='10'>"
            + "<feOffset in='SourceGraphic' dx='1'/></filter></defs>"
            + "<rect x='12' y='2' width='4' height='3' fill='lime' transform='scale(2)' filter='url(#f)'/></svg>";
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(ReadFilterGraph(svg));
        Assert.True(raster.GetPixel(17, 5).G > 240);
        Assert.Equal(0, raster.GetPixel(14, 5).A);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SvgFilterGraphUsesOriginalGeometryForEnclosingObjectBounds(bool routed) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 20'><defs>"
            + "<filter id='inner' filterUnits='userSpaceOnUse' x='0' y='0' width='20' height='20'>"
            + "<feOffset " + (routed ? "in='SourceGraphic' " : "") + "dx='3'/></filter>"
            + "<filter id='outer'><feColorMatrix values='0 0 0 0 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1'/></filter></defs>"
            + "<g filter='url(#outer)'><rect x='4' y='4' width='4' height='4' fill='lime' filter='url(#inner)'/></g></svg>";
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(ReadFilterGraph(svg));
        Assert.True(raster.GetPixel(5, 5).R > 240);
        Assert.Equal(0, raster.GetPixel(1, 5).A);
        Assert.Equal(0, raster.GetPixel(10, 5).A);
    }

    [Fact]
    public void SvgFilterGraphDisablesEmptyFractionalRegionWithoutLosingTheDocument() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><defs>"
            + "<filter id='f' filterUnits='userSpaceOnUse' x='2.3' y='2' width='0' height='4'><feColorMatrix/></filter></defs>"
            + "<rect x='2' y='2' width='4' height='4' fill='red' filter='url(#f)'/>"
            + "<rect x='7' y='7' width='2' height='2' fill='lime'/></svg>";
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(ReadFilterGraph(svg));
        Assert.Equal(0, raster.GetPixel(3, 3).A);
        Assert.True(raster.GetPixel(8, 8).G > 240);
    }

    [Fact]
    public void SvgFilterGraphResolvesImplicitInputsAndClosestPrecedingResult() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 10'><defs>"
            + "<filter id='f' filterUnits='userSpaceOnUse' x='0' y='0' width='20' height='10'>"
            + "<feOffset in='SourceGraphic' dx='2' result='same'/><feOffset dx='2' result='same'/>"
            + "<feOffset in='same' dx='2'/></filter></defs>"
            + "<rect x='2' y='2' width='3' height='3' fill='red' filter='url(#f)'/></svg>";
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(ReadFilterGraph(svg));
        Assert.True(raster.GetPixel(9, 3).R > 240);
        Assert.Equal(0, raster.GetPixel(3, 3).A);
        Assert.Equal(0, raster.GetPixel(6, 3).A);
    }

    [Fact]
    public void SvgFilterGraphKeepsLinkGeometryIndependentOfFilterImageRegion() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 10'><defs>"
            + "<filter id='f'><feOffset in='SourceGraphic' dx='1'/></filter></defs>"
            + "<a href='https://example.test/filtered' filter='url(#f)'><rect x='4' y='2' width='3' height='4' fill='red'/></a></svg>";
        OfficeDrawing drawing = ReadFilterGraph(svg);
        OfficeDrawingLink link = Assert.Single(FilterGraphElements(drawing).OfType<OfficeDrawingLink>());
        Assert.Equal(4D, link.X); Assert.Equal(2D, link.Y);
        Assert.Equal(3D, link.Width); Assert.Equal(4D, link.Height);
        Assert.True(OfficeDrawingRasterRenderer.Render(drawing).GetPixel(6, 3).R > 240);
    }

    [Theory]
    [InlineData("<feOffset in='later' dx='1' result='a'/><feOffset in='a' result='later'/>")]
    [InlineData("<feColorMatrix values='1 0'/>")]
    [InlineData("<feBlend in='SourceGraphic' in2='SourceAlpha' mode='hue'/>")]
    [InlineData("<feOffset in='SourceGraphic'/><feColorMatrix color-interpolation-filters='sRGB'/>")]
    [InlineData("<feBlend in='SourceGraphic' in2='SourceAlpha' no-composite='no-composite'/>")]
    public void SvgFilterGraphReportsUnsupportedGraphAndPreservesSource(string primitives) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><defs><filter id='f'>"
            + primitives + "</filter></defs><rect x='2' y='2' width='5' height='5' fill='lime' filter='url(#f)'/></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(1, unsupported);
        Assert.True(OfficeDrawingRasterRenderer.Render(drawing!).GetPixel(4, 4).G > 240);
    }

    [Fact]
    public void SvgFilterGraphDoesNotRasterizeSearchableTextOrNestedLinks() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 30'><defs>"
            + "<filter id='f'><feColorMatrix/></filter></defs><g filter='url(#f)'>"
            + "<text x='5' y='15'>Searchable label</text><a href='https://example.test/kept'>"
            + "<rect x='60' y='2' width='5' height='5'/></a></g></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(1, unsupported);
        Assert.Contains("Searchable label", OfficeDrawingSvgExporter.ToSvg(drawing!), StringComparison.Ordinal);
        Assert.Single(FilterGraphElements(drawing!).OfType<OfficeDrawingLink>());
        Assert.Empty(FilterGraphElements(drawing!).OfType<OfficeDrawingImage>());
    }

    [Fact]
    public void SvgFilterGraphBoundsAllocatedRegionAndSharesWorkAcrossConsumers() {
        const string huge = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><defs>"
            + "<filter id='f' filterUnits='userSpaceOnUse' width='100000' height='100000'><feColorMatrix/></filter></defs>"
            + "<rect width='10' height='10' fill='red' filter='url(#f)'/></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(huge), out OfficeDrawing? bounded, out int unsupported));
        Assert.Equal(1, unsupported);
        Assert.Empty(FilterGraphElements(bounded!).OfType<OfficeDrawingImage>());
        const string shared = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 400 400'><defs>"
            + "<filter id='f'><feGaussianBlur in='SourceAlpha' stdDeviation='32'/></filter></defs>"
            + "<rect x='40' y='40' width='300' height='300' filter='url(#f)'/>"
            + "<rect x='40' y='40' width='300' height='300' filter='url(#f)'/></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(shared), out OfficeDrawing? drawing, out unsupported));
        Assert.Equal(1, unsupported);
        Assert.Single(FilterGraphElements(drawing!).OfType<OfficeDrawingImage>());
        Assert.Single(FilterGraphElements(drawing!).OfType<OfficeDrawingShape>());
    }

    [Fact]
    public void SvgFilterGraphObservesCancellationAfterSourceImport() {
        using var cancellation = new CancellationTokenSource();
        var options = new OfficeSvgDrawingReaderOptions {
            CancellationToken = cancellation.Token,
            ForeignObjectRenderer = context => {
                var content = new OfficeDrawing(context.Width, context.Height);
                content.AddShape(OfficeShape.Rectangle(context.Width, context.Height), 0D, 0D);
                cancellation.Cancel();
                return content;
            }
        };
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 20'><defs>"
            + "<filter id='f'><feGaussianBlur in='SourceAlpha' stdDeviation='2'/></filter></defs>"
            + "<g filter='url(#f)'><foreignObject width='10' height='10'><div xmlns='http://www.w3.org/1999/xhtml'>source</div></foreignObject></g></svg>";
        Assert.Throws<OperationCanceledException>(() => OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), options, out _, out _));
    }

    private static OfficeDrawing ReadFilterGraph(string svg) {
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        return drawing!;
    }

    private static IEnumerable<OfficeDrawingElement> FilterGraphElements(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            yield return element;
            OfficeDrawing? inner = element is OfficeDrawingGroup group ? group.Drawing
                : element is OfficeDrawingEffectGroup effect ? effect.Drawing : null;
            if (inner != null) foreach (OfficeDrawingElement child in FilterGraphElements(inner)) yield return child;
        }
    }
}
