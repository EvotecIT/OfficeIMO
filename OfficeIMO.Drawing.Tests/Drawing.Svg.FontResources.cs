using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingSvgFontResourceTests {
    [Theory]
    [InlineData("<text x='8' y='24' textLength='32' lengthAdjust='spacingAndGlyphs'>A</text>")]
    [InlineData("<text x='8' y='24' transform='translate(3 0)'>A</text>")]
    [InlineData("<g style='mix-blend-mode:multiply'><text x='8' y='24'>A</text></g>")]
    [InlineData("<a href='https://example.com'><text x='8' y='24'>A</text></a>")]
    [InlineData("<text x='8' y='24' style='mix-blend-mode:multiply'>A</text>")]
    [InlineData("<defs><mask id='m' maskUnits='userSpaceOnUse' x='0' y='0' width='100' height='40'><text x='8' y='24' fill='white'>A</text></mask></defs><rect width='100' height='40' mask='url(#m)'/>")]
    [InlineData("<defs><pattern id='p' patternUnits='userSpaceOnUse' width='20' height='30'><text x='0' y='24'>A</text></pattern></defs><rect width='100' height='40' fill='url(#p)'/>")]
    [InlineData("<defs><pattern id='p' patternUnits='userSpaceOnUse' width='20' height='30'><text x='0' y='24'>A</text></pattern></defs><path d='M4 20 L40 10 L80 20' fill='none' stroke='url(#p)' stroke-width='8'/>")]
    [InlineData("<defs><marker id='m' markerUnits='userSpaceOnUse' markerWidth='20' markerHeight='30'><text x='0' y='24'>A</text></marker></defs><line x1='10' y1='5' x2='60' y2='5' stroke='black' marker-start='url(#m)'/>")]
    [InlineData("<defs><symbol id='s' viewBox='0 0 100 40'><text x='8' y='24'>A</text></symbol></defs><use href='#s' width='100' height='40'/>")]
    public void SvgTextEffectsRetainCallerSuppliedFontResources(string content) {
        var options = new OfficeSvgDrawingReaderOptions();
        options.Fonts.Add("FixtureFont", ManagedTextShapingTestAssets.CreateFont('A'));
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='100' height='40' font-family='FixtureFont' font-size='20'>" + content + "</svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), options, out OfficeDrawing? drawing, out int unsupported));
        Assert.True(unsupported == 0, "Unsupported content: " + content);
        Assert.NotNull(drawing);
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        drawing!.AppendFontDiagnostics(diagnostics);
        Assert.DoesNotContain(diagnostics, item => item.Code == OfficeImageExportDiagnosticCodes.FontSubstituted);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.True(raster.GetPixels().Any(value => value != 0));
        static void CheckFonts(OfficeDrawing scene) {
            foreach (OfficeDrawingElement element in scene.Elements) {
                if (element is OfficeDrawingText text) Assert.Null(scene.Fonts.CreateSubstitutionDiagnostic(text.Text, text.Font.FamilyName));
                if (element is OfficeDrawingGroup group) CheckFonts(group.Drawing);
                if (element is OfficeDrawingEffectGroup effect) CheckFonts(effect.Drawing);
            }
        }
        CheckFonts(drawing);
    }
}
