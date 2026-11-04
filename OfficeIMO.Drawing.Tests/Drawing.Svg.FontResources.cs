using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingSvgFontResourceTests {
#if NET8_0_OR_GREATER
    [Theory]
    [InlineData("normal")]
    [InlineData("100 900")]
    public void RealWoff2FontFaceDescriptorsPermitBoundedRasterFallback(string weight) {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "OpenSans-Regular.woff2"));
        byte[] bytes = FontSafetySvg(SafetyFontFace(font, ";font-style:normal;font-weight:" + weight, "woff2"));

        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(bytes));
        Assert.True(OfficeSvgDrawingReader.TryRead(bytes, out OfficeDrawing? drawing, out int unsupported));
        if (weight == "100 900") {
            // The editable importer reports the range as unsupported; fallback must use the original SVG.
            Assert.True(unsupported > 0);
        } else {
            Assert.Equal(0, unsupported);
            Assert.Equal(OfficeFontContainerFormat.Woff2, Assert.Single(drawing!.Fonts.Faces).ContainerFormat);
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            drawing.AppendFontDiagnostics(diagnostics);
            Assert.DoesNotContain(diagnostics, item => item.Code == OfficeImageExportDiagnosticCodes.FontSubstituted);
            Assert.Contains(OfficeDrawingRasterRenderer.Render(drawing).GetPixels(), value => value != 0);
        }
    }
#endif

    [Theory]
    [InlineData("{FACE}.paint{font-family:FixtureFont}")]
    [InlineData("{FACE}.paint{f\\6f nt-family:FixtureFont}")]
    [InlineData("{FACE}\\40 font-face{font-weight:900}")]
    [InlineData("{FACE}@media all{.paint{font-family:FixtureFont}}")]
    [InlineData("{FACE}.paint{fill:red;.nested{stroke-width:4096}}")]
    [InlineData("{FACE}.paint{fill:'red}")]
    [InlineData("{FACE}.paint{--x:([)];fill:red}")]
    [InlineData("{FACE}/* unterminated")]
    public void ValidFontFaceDoesNotExemptElementGeometryOrMalformedStylesheets(string stylesheet) {
        string face = SafetyFontFace(ManagedTextShapingTestAssets.CreateFont('A'));
        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(FontSafetySvg(face)));

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(FontSafetySvg(stylesheet.Replace("{FACE}", face))));
    }

    [Theory]
    [InlineData(";font-weight:900 100")]
    [InlineData(";font-weight:0 900")]
    [InlineData(";font-weight:100 1001")]
    [InlineData(";font-weight:100 400 900")]
    [InlineData(";font-weight:100 900!important")]
    [InlineData(";font-style:oblique 90deg")]
    [InlineData(";broken-descriptor")]
    [InlineData(";transform:scale(4096)")]
    [InlineData(";stroke-width:4096")]
    public void RasterSafetyRequiresValidFontFaceDescriptors(string descriptors) {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A');
        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(FontSafetySvg(SafetyFontFace(font))));

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(FontSafetySvg(SafetyFontFace(font, descriptors))));
    }

    [Fact]
    public void RasterSafetyRejectsFontDecodedSizeExpansionBeforeFallback() {
        byte[] woff = ManagedTextShapingTestAssets.CreateWoff(ManagedTextShapingTestAssets.CreateFont('A'));
        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(FontSafetySvg(SafetyFontFace(woff, format: "woff"))));
        uint declared = 8U * 1024U * 1024U + 1U;
        for (int index = 0; index < 4; index++) woff[16 + index] = (byte)(declared >> (24 - index * 8));

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(FontSafetySvg(SafetyFontFace(woff, format: "woff"))));
    }

    [Fact]
    public void RasterSafetySharesFontFaceBudgetAcrossStylesheets() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A');
        string faces = string.Concat(Enumerable.Repeat(SafetyFontFace(font), 128));
        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(FontSafetySvg(faces)));

        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(
            FontSafetySvg(faces, "<style>" + SafetyFontFace(font) + "</style>")));
    }

    private static string SafetyFontFace(byte[] font, string descriptors = "", string format = "truetype") {
        string mime = format == "truetype" ? "ttf" : format;
        return "@font-face{font-family:FixtureFont;src:url('data:font/" + mime + ";base64,"
            + Convert.ToBase64String(font) + "') format('" + format + "')" + descriptors + "}";
    }

    private static byte[] FontSafetySvg(string stylesheet, string content = "") =>
        Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='320' height='60' viewBox='0 0 320 60'>"
            + "<style>" + stylesheet + "</style><text x='2' y='30' font-family='FixtureFont' font-size='12'>A</text>"
            + content + "</svg>");

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
        Assert.Contains(raster.GetPixels(), value => value != 0);
        static void CheckFonts(OfficeDrawing scene) {
            foreach (OfficeDrawingElement element in scene.Elements) {
                if (element is OfficeDrawingText text) Assert.Null(scene.Fonts.CreateSubstitutionDiagnostic(text.Text, text.Font.FamilyName));
                if (element is OfficeDrawingGroup group) CheckFonts(group.Drawing);
                if (element is OfficeDrawingEffectGroup effect) CheckFonts(effect.Drawing);
            }
        }
        CheckFonts(drawing);
    }

    [Fact]
    public void FittedRootViewportRetainsCallerSuppliedFontResources() {
        var options = new OfficeSvgDrawingReaderOptions();
        options.Fonts.Add("FixtureFont", ManagedTextShapingTestAssets.CreateFont('A'));
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='200' height='80' viewBox='0 0 100 40' "
            + "font-family='FixtureFont' font-size='20'><text x='8' y='24'>A</text></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            Encoding.UTF8.GetBytes(svg), options, out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Equal("FixtureFont", Assert.Single(drawing!.Fonts.Faces).FamilyName);

        var diagnostics = new List<OfficeImageExportDiagnostic>();
        drawing.AppendFontDiagnostics(diagnostics);
        Assert.DoesNotContain(diagnostics, item => item.Code == OfficeImageExportDiagnosticCodes.FontSubstituted);
        Assert.Contains("@font-face", OfficeDrawingSvgExporter.ToSvg(drawing), StringComparison.Ordinal);
    }
}
