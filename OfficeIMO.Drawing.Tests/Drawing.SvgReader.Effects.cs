using System.Text;
using System.Diagnostics;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeSvgDrawingReader_MapsLocalMasksAndBlendModesIntoSharedEffectGroups() {
        const string svg = """
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 4">
              <defs>
                <mask id="left-half" maskUnits="userSpaceOnUse" maskContentUnits="userSpaceOnUse" x="0" y="0" width="5" height="4" style="mask-type:luminance">
                  <rect x="0" y="0" width="10" height="4" fill="white" />
                </mask>
              </defs>
              <rect x="0" y="0" width="10" height="4" fill="blue" />
              <g mask="url(#left-half)" style="mix-blend-mode:multiply">
                <rect x="0" y="0" width="10" height="4" fill="red" />
              </g>
            </svg>
            """;

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg),
            out OfficeDrawing? drawing,
            out int unsupported));

        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeDrawingEffectGroup effect = Assert.Single(drawing!.Elements.OfType<OfficeDrawingEffectGroup>());
        Assert.Equal(OfficeBlendMode.Multiply, effect.BlendMode);
        Assert.NotNull(effect.SoftMask);
        Assert.Equal(OfficeSoftMaskMode.Luminosity, effect.SoftMask!.Mode);

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        OfficeColor masked = raster.GetPixel(2, 2);
        OfficeColor outside = raster.GetPixel(8, 2);
        Assert.True(masked.R < 10 && masked.G < 10 && masked.B < 10, masked.ToString());
        Assert.True(outside.B > 240 && outside.R < 10, outside.ToString());

        string roundTrip = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains("mix-blend-mode:multiply", roundTrip, StringComparison.Ordinal);
        Assert.Contains("<mask id=\"officeimo-mask-", roundTrip, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeSvgDrawingReader_DiagnosesObjectBoundingBoxMasksInsteadOfApplyingWrongCoordinates() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 4'><defs>"
            + "<mask id='default-units'><rect width='1' height='1' fill='white'/></mask></defs>"
            + "<rect width='10' height='4' fill='red' mask='url(#default-units)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg),
            out OfficeDrawing? drawing,
            out int unsupported));

        Assert.NotNull(drawing);
        Assert.Equal(1, unsupported);
        OfficeDrawingEffectGroup effect = Assert.Single(drawing!.Elements.OfType<OfficeDrawingEffectGroup>());
        Assert.Null(effect.SoftMask);
        Assert.Single(effect.Drawing.Shapes);
    }

    [Fact]
    public void OfficeSvgDrawingReader_TransformsUserSpaceMaskRegionWithReferencingGroup() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 4'><defs>"
            + "<mask id='translated' maskUnits='userSpaceOnUse' x='0' y='0' width='2' height='4'>"
            + "<rect width='2' height='4' fill='white'/></mask></defs>"
            + "<rect width='10' height='4' fill='blue'/>"
            + "<g transform='translate(4 0)' mask='url(#translated)'><rect width='2' height='4' fill='red'/></g></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg),
            out OfficeDrawing? drawing,
            out int unsupported));

        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing!);
        Assert.True(raster.GetPixel(1, 2).B > 240);
        Assert.True(raster.GetPixel(5, 2).R > 240);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" x='-10%'")]
    public void OfficeSvgDrawingReader_UserSpaceMaskPercentagesUseTheViewBoxOrigin(string xAttribute) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='100 0 10 4'><defs>"
            + "<mask id='default-region' maskUnits='userSpaceOnUse'" + xAttribute + "><rect x='100' width='10' height='4' fill='white'/></mask></defs>"
            + "<rect x='100' width='10' height='4' fill='red' mask='url(#default-region)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg),
            out OfficeDrawing? drawing,
            out int unsupported));

        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing!);
        Assert.True(raster.GetPixel(5, 2).R > 240, raster.GetPixel(5, 2).ToString());
    }

    [Fact]
    public void OfficeSvgDrawingReader_MaskContentInheritsPaintFromDefinitionAncestors() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 4' fill='white'><defs>"
            + "<mask id='inherited-paint' maskUnits='userSpaceOnUse' x='0' y='0' width='10' height='4'>"
            + "<rect width='10' height='4'/></mask></defs>"
            + "<rect width='10' height='4' fill='red' mask='url(#inherited-paint)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg),
            out OfficeDrawing? drawing,
            out int unsupported));

        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing!);
        OfficeColor visible = raster.GetPixel(5, 2);
        Assert.True(visible.R > 240 && visible.G < 10 && visible.B < 10, visible.ToString());
    }

    [Fact]
    public void OfficeSvgDrawingReader_DoesNotCreateEffectGroupsForNormalBlendMode() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 4'>"
            + "<g style='mix-blend-mode:normal'><rect width='10' height='4' fill='red'/></g></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg),
            out OfficeDrawing? drawing,
            out int unsupported));

        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        Assert.Empty(drawing!.Elements.OfType<OfficeDrawingEffectGroup>());
        Assert.Single(drawing.Shapes);
    }

    [Fact]
    public void OfficeSvgDrawingReader_DiagnosesUnsupportedFiltersWithoutDroppingSupportedGeometry() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><rect width='10' height='10' fill='red' filter='url(#blur)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg),
            out OfficeDrawing? drawing,
            out int unsupported));

        Assert.NotNull(drawing);
        Assert.Single(drawing!.Shapes);
        Assert.Equal(1, unsupported);
    }

    [Fact]
    public void OfficeSvgDrawingReader_RendersBoundedDropShadowFiltersAsVectorEffectGroups() {
        const string svg = """
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 12 10">
              <defs>
                <filter id="shadow">
                  <feDropShadow dx="2" dy="1" stdDeviation="1" flood-color="red" flood-opacity="0.6" />
                </filter>
              </defs>
              <rect x="2" y="2" width="4" height="4" fill="blue" filter="url(#shadow)" />
            </svg>
            """;

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg),
            out OfficeDrawing? drawing,
            out int unsupported));

        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeDrawingEffectGroup outer = Assert.Single(drawing!.Elements.OfType<OfficeDrawingEffectGroup>());
        OfficeDrawing filtered = outer.Drawing;
        Assert.Equal(2, filtered.Elements.OfType<OfficeDrawingEffectGroup>().Count());
        OfficeDrawing blurSamples = filtered.Elements.OfType<OfficeDrawingEffectGroup>().First().Drawing;
        Assert.Equal(9, blurSamples.Elements.OfType<OfficeDrawingEffectGroup>().Count());

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.True(raster.GetPixel(3, 3).B > 200, raster.GetPixel(3, 3).ToString());
        Assert.True(raster.GetPixel(8, 5).R > raster.GetPixel(8, 5).B, raster.GetPixel(8, 5).ToString());
        string roundTrip = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains("fill=\"#FF0000\"", roundTrip, StringComparison.Ordinal);
        Assert.Contains("fill=\"#0000FF\"", roundTrip, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeSvgDrawingReader_AppliesStaticFiltersDeclaredOnTheRootViewport() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 12 10' filter='url(#shadow)'>"
            + "<defs><filter id='shadow'><feDropShadow dx='2' dy='1' stdDeviation='0' flood-color='red'/></filter></defs>"
            + "<rect x='2' y='2' width='4' height='4' fill='blue'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg),
            out OfficeDrawing? drawing,
            out int unsupported));

        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        Assert.Single(drawing!.Elements.OfType<OfficeDrawingEffectGroup>());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.True(raster.GetPixel(3, 3).B > 200);
        Assert.True(raster.GetPixel(7, 6).R > raster.GetPixel(7, 6).B);
    }

    [Fact]
    public void OfficeSvgDrawingReader_ComposesGaussianBlurAndOffsetWithoutRasterImages() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 8'><defs>"
            + "<filter id='moved'><feGaussianBlur stdDeviation='0'/><feOffset dx='2' dy='1'/></filter></defs>"
            + "<rect x='1' y='1' width='3' height='3' fill='lime' filter='url(#moved)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg),
            out OfficeDrawing? drawing,
            out int unsupported));

        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        Assert.Empty(drawing!.Images);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Equal(0, raster.GetPixel(1, 1).A);
        Assert.True(raster.GetPixel(3, 2).G > 200, raster.GetPixel(3, 2).ToString());
    }

    [Fact]
    public void OfficeSvgDrawingReader_DiagnosesFilterGraphsOutsideTheStaticPrimitiveSubset() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><defs>"
            + "<filter id='graph'><feGaussianBlur stdDeviation='1' result='blur'/><feBlend in='SourceGraphic' in2='blur'/></filter></defs>"
            + "<rect width='10' height='10' fill='red' filter='url(#graph)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg),
            out OfficeDrawing? drawing,
            out int unsupported));

        Assert.NotNull(drawing);
        Assert.Equal(1, unsupported);
        Assert.Single(drawing!.Shapes);
    }

    [Fact]
    public void OfficeSvgDrawingReader_UsesBoundedForeignObjectRendererWithClipAndTransform() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 8'>"
            + "<foreignObject x='2' y='1' width='4' height='3' transform='translate(1 1)'>"
            + "<div xmlns='http://www.w3.org/1999/xhtml' style='color:red'>Foreign</div>"
            + "</foreignObject></svg>";
        OfficeSvgForeignObjectContext? request = null;
        var options = new OfficeSvgDrawingReaderOptions {
            ForeignObjectRenderer = context => {
                request = context;
                var nested = new OfficeDrawing(context.Width, context.Height);
                var shape = OfficeShape.Rectangle(context.Width, context.Height);
                shape.FillColor = OfficeColor.Red;
                nested.AddShape(shape, 0D, 0D);
                return nested;
            }
        };

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg),
            options,
            out OfficeDrawing? drawing,
            out int unsupported));

        Assert.NotNull(drawing);
        Assert.NotNull(request);
        Assert.Equal(4D, request!.Width);
        Assert.Equal(3D, request.Height);
        Assert.Contains("Foreign", request.Html, StringComparison.Ordinal);
        Assert.Equal(0, unsupported);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing!);
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(2, 1));
        Assert.True(raster.GetPixel(4, 3).R > 240);
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(7, 5));
    }

    [Fact]
    public void OfficeSvgDrawingReader_DiagnosesForeignObjectRendererViewportMismatch() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 8'>"
            + "<foreignObject width='4' height='3'><div xmlns='http://www.w3.org/1999/xhtml'>Mismatch</div></foreignObject></svg>";
        var options = new OfficeSvgDrawingReaderOptions {
            ForeignObjectRenderer = _ => new OfficeDrawing(3D, 3D)
        };

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg),
            options,
            out OfficeDrawing? drawing,
            out int unsupported));

        Assert.NotNull(drawing);
        Assert.Empty(drawing!.Elements);
        Assert.Equal(1, unsupported);
    }

    [Fact]
    public void OfficeSvgDrawingReader_DiagnosesUnsupportedForeignObjectEffectsWithoutDroppingContent() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 8'><defs>"
            + "<filter id='shadow'><feDropShadow dx='1' dy='1' stdDeviation='1'/></filter></defs>"
            + "<foreignObject width='4' height='3' filter='url(#shadow)' style='mix-blend-mode:multiply'>"
            + "<div xmlns='http://www.w3.org/1999/xhtml'>Foreign</div></foreignObject></svg>";
        var options = new OfficeSvgDrawingReaderOptions {
            ForeignObjectRenderer = context => {
                var nested = new OfficeDrawing(context.Width, context.Height);
                nested.AddShape(OfficeShape.Rectangle(context.Width, context.Height), 0D, 0D);
                return nested;
            }
        };

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg),
            options,
            out OfficeDrawing? drawing,
            out int unsupported));

        Assert.NotNull(drawing);
        Assert.NotEmpty(drawing!.Elements);
        Assert.Equal(1, unsupported);
    }

    [Fact]
    public void OfficeSvgDrawingReader_AppliesMasksToTextAndUseElements() {
        const string prefix = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 30'><defs>"
            + "<mask id='m' maskUnits='userSpaceOnUse' x='0' y='0' width='20' height='30'><rect width='20' height='30' fill='white'/></mask>"
            + "<rect id='shape' width='4' height='4' fill='red'/></defs>";
        string textSvg = prefix + "<text x='1' y='20' mask='url(#m)'>A</text></svg>";
        string useSvg = prefix + "<use href='#shape' x='10' y='2' mask='url(#m)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(textSvg),
            out OfficeDrawing? textDrawing,
            out int textUnsupported));
        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(useSvg),
            out OfficeDrawing? useDrawing,
            out int useUnsupported));

        Assert.NotNull(textDrawing);
        Assert.NotNull(useDrawing);
        Assert.Equal(0, textUnsupported);
        Assert.Equal(0, useUnsupported);
        Assert.NotNull(Assert.Single(textDrawing!.Elements.OfType<OfficeDrawingEffectGroup>()).SoftMask);
        Assert.NotNull(Assert.Single(useDrawing!.Elements.OfType<OfficeDrawingEffectGroup>()).SoftMask);
    }

    [Fact]
    public void OfficeSvgDrawingReader_BoundsCyclicMaskReferencesAndStaysDeterministic() {
        const string svg = """
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10">
              <defs><mask id="cycle"><g mask="url(#cycle)"><rect width="10" height="10" fill="white"/></g></mask></defs>
              <rect width="10" height="10" fill="red" mask="url(#cycle)"/>
            </svg>
            """;
        byte[] bytes = Encoding.UTF8.GetBytes(svg);

        Assert.True(OfficeSvgDrawingReader.TryRead(bytes, out OfficeDrawing? first, out int firstUnsupported));
        Assert.True(OfficeSvgDrawingReader.TryRead(bytes, out OfficeDrawing? second, out int secondUnsupported));

        Assert.NotNull(first);
        Assert.NotNull(second);
        Assert.True(firstUnsupported > 0);
        Assert.Equal(firstUnsupported, secondUnsupported);
        Assert.Equal(OfficeDrawingSvgExporter.ToSvg(first!), OfficeDrawingSvgExporter.ToSvg(second!));
    }

    [Fact]
    public void OfficeSvgDrawingReader_DeterministicMutationFuzzRemainsBounded() {
        byte[] seed = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 20'><defs><mask id='m' maskUnits='userSpaceOnUse'><rect width='20' height='20' fill='white'/></mask></defs><g mask='url(#m)' style='mix-blend-mode:multiply'><path d='M0 0 L40 20 Z' fill='red'/></g></svg>");
        uint state = 0x5A17C0DEU;
        var timer = Stopwatch.StartNew();

        for (int caseNumber = 0; caseNumber < 64; caseNumber++) {
            byte[] candidate = (byte[])seed.Clone();
            for (int mutation = 0; mutation < 8; mutation++) {
                state = unchecked(state * 1664525U + 1013904223U);
                int index = (int)(state % (uint)candidate.Length);
                state = unchecked(state * 1664525U + 1013904223U);
                candidate[index] = (byte)(32 + state % 95U);
            }

            bool firstSucceeded = OfficeSvgDrawingReader.TryRead(candidate, out OfficeDrawing? first, out int firstUnsupported);
            bool secondSucceeded = OfficeSvgDrawingReader.TryRead(candidate, out OfficeDrawing? second, out int secondUnsupported);

            Assert.Equal(firstSucceeded, secondSucceeded);
            Assert.Equal(firstUnsupported, secondUnsupported);
            if (firstSucceeded && first != null && second != null) {
                Assert.Equal(OfficeDrawingSvgExporter.ToSvg(first), OfficeDrawingSvgExporter.ToSvg(second));
            }
        }

        Assert.True(timer.Elapsed < TimeSpan.FromSeconds(10D), "SVG mutation fuzz pass exceeded its test budget: " + timer.Elapsed + ".");
    }
}
