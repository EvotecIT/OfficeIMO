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
