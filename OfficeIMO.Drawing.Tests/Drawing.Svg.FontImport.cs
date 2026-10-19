using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.ContentSafety;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingSvgFontImportTests {
    [Theory]
    [InlineData("\"SVG Fixture, \\\"Sans\\\" & Co\"")]
    [InlineData("\"SVG Fixture\\2C  \\22 Sans\\22  \\26  Co\"")]
    public void SvgFontFamilyEscapesResolveWithoutFalseSubstitutionDiagnostics(string request) {
        var fonts = new OfficeFontFaceCollection();
        fonts.Add("SVG Fixture, \"Sans\" & Co", ManagedTextShapingTestAssets.CreateFont('A'));

        Assert.True(fonts.TryResolveFaceForText("A", request, OfficeFontStyle.Regular, out OfficeFontFace? face));
        Assert.Equal("SVG Fixture, \"Sans\" & Co", face!.FamilyName);
        Assert.Null(fonts.CreateSubstitutionDiagnostic("A", request));
    }

    [Theory]
    [InlineData("")]
    [InlineData("page-1-")]
    public void ExportedPortableFontReopensWithTheSamePaintAsExplicitFontRegistration(string prefix) {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Carlito-Regular.ttf"));
        const string family = "SVG Fixture, \"Sans\" & Co";
        var descriptor = new OfficeFontFaceDescriptor(500, 87.5D, OfficeFontSlant.Oblique, 10D);
        Assert.True(OfficeFontUnicodeRangeSet.TryParseCss("U+0041-005A", out OfficeFontUnicodeRangeSet? ranges));
        var original = new OfficeDrawing(180, 60);
        original.Fonts.Add(family, font, descriptor, ranges!);
        original.AddText("ABC", 8, 4, 160, 48, new OfficeFontInfo("\"SVG Fixture, \\\"Sans\\\" & Co\"", 24, descriptor));
        string exported = OfficeDrawingSvgExporter.ToSvg(original, 1D, OfficeSvgSizeUnit.Pixel, null, prefix);
        XElement xml = XElement.Parse(exported);
        Assert.Contains("@font-face", xml.Descendants().Single(element => element.Name.LocalName == "style").Value);
        string textFamily = xml.Descendants().First(element => element.Name.LocalName == "text").Attribute("font-family")!.Value;

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(exported), out OfficeDrawing? reopened, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.True(reopened!.Fonts.TryResolveFaceForText("ABC", textFamily, descriptor, out OfficeFontFace? imported));
        Assert.Equal(descriptor, imported!.Descriptor);
        Assert.Equal(font, imported.Data);
        Assert.Contains(reopened.Fonts.Faces, face => face.FamilyName == prefix + family && face.UnicodeRanges.Equals(ranges));

        // Independent control of import wiring: identical exported geometry, with the embedded
        // stylesheet removed and the already supported explicit font-registration path instead.
        xml.Descendants().Where(element => element.Name.LocalName == "style").Remove();
        var options = new OfficeSvgDrawingReaderOptions();
        options.Fonts.Add(prefix + original.Fonts.Faces[0].ResourceFamilyName, font, descriptor);
        options.Fonts.Add(prefix + family, font, descriptor, ranges!);
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(xml.ToString()), options, out OfficeDrawing? control, out int controlUnsupported));
        Assert.Equal(0, controlUnsupported);
        byte[] pixels = OfficeDrawingRasterRenderer.Render(reopened).GetPixels();
        Assert.Contains(pixels, channel => channel != 0);
        Assert.Equal(OfficeDrawingRasterRenderer.Render(control!).GetPixels(), pixels);
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        reopened.AppendFontDiagnostics(diagnostics);
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.FontSubstituted);
        Assert.Contains("@font-face", OfficeDrawingSvgExporter.ToSvg(reopened));
    }

    [Theory]
    [InlineData("", "<text x='8' y='24' transform='translate(1 0)'>A</text>")]
    [InlineData("opacity='.5'", "<text x='8' y='24'>A</text>")]
    [InlineData("width='200' height='80'", "<text x='8' y='24'>A</text>")]
    public void EmbeddedSvgFontScopeReachesPaintedTextEffectsAndFittedViewports(string rootAttributes, string content) {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A');
        byte[] bytes = SvgFontImportDocument(FontFace("FixtureFont", font), content, rootAttributes);
        var options = new OfficeSvgDrawingReaderOptions();
        options.Fonts.Add("CallerFont", font);

        Assert.True(OfficeSvgDrawingReader.TryRead(bytes, options, out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Single(options.Fonts.Faces);
        Assert.Equal("CallerFont", options.Fonts.Faces[0].FamilyName);
        Assert.Contains(drawing!.Fonts.Faces, face => face.FamilyName == "FixtureFont");
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        drawing.AppendFontDiagnostics(diagnostics);
        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.FontSubstituted);
        Assert.Contains(OfficeDrawingRasterRenderer.Render(drawing).GetPixels(), channel => channel != 0);
    }

    [Fact]
    public void EmbeddedSvgFontIsAvailableToStrictContentSafetyTextInspection() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A');
        byte[] bytes = SvgFontImportDocument(FontFace("FixtureFont", font),
            "<rect width='100' height='40' fill='white'/><text x='8' y='24' fill='white'>A</text>");

        OfficeContentSafetyReport report = OfficeSvgDrawingReader.InspectContentSafety(bytes);

        Assert.Contains(report.Findings, finding => finding.TextPreview == "A");
    }

    [Fact]
    public void EmbeddedSvgFontDescriptorsSelectSeparateWeightsAndUnicodeRanges() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A', 'B');
        string sources = FontFace("FixtureFont", font, ";font-weight:300;unicode-range:U+0041")
            + FontFace("FixtureFont", font, ";font-weight:500;unicode-range:U+0041")
            + FontFace("FixtureFont", font, ";font-weight:500;unicode-range:U+0042");
        Assert.True(OfficeSvgDrawingReader.TryRead(SvgFontImportDocument(sources), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Equal(3, drawing!.Fonts.Faces.Count);
        Assert.True(drawing.Fonts.TryResolveFaceForText("A", "FixtureFont", new OfficeFontFaceDescriptor(450), out OfficeFontFace? a));
        Assert.Equal(500, a!.Descriptor.Weight);
        Assert.False(a.UnicodeRanges.Contains('B'));
        Assert.True(drawing.Fonts.TryResolveFaceForText("B", "FixtureFont", new OfficeFontFaceDescriptor(500), out OfficeFontFace? b));
        Assert.False(b!.UnicodeRanges.Contains('A'));
    }

    [Fact]
    public void EmbeddedSvgFontUsesTheCanonicalWoffDecoderForQuotedCssSources() {
        byte[] woff = ManagedTextShapingTestAssets.CreateWoff(ManagedTextShapingTestAssets.CreateFont('A'));
        byte[] bytes = SvgFontImportDocument("@font-face{font-family:FixtureFont;src:url('data:font/woff;base64,"
            + Convert.ToBase64String(woff) + "') format('woff');}", "<text x='8' y='24'>A</text>");

        Assert.True(OfficeSvgDrawingReader.TryRead(bytes, out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.True(drawing!.Fonts.TryResolveFaceForText("A", "FixtureFont", OfficeFontStyle.Regular, out OfficeFontFace? face));
        Assert.Equal(OfficeFontContainerFormat.Woff, face!.ContainerFormat);
        Assert.Contains(OfficeDrawingRasterRenderer.Render(drawing).GetPixels(), channel => channel != 0);
    }

    private static string FontFace(string family, byte[] data, string descriptors = "") =>
        "@font-face{font-family:'" + family + "';src:url(data:font/ttf;base64," + Convert.ToBase64String(data) + ") format('truetype')" + descriptors + ";}";

    private static byte[] SvgFontImportDocument(string css, string content = "", string rootAttributes = "") =>
        Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 40' font-family='FixtureFont' font-size='20' "
            + rootAttributes + "><defs><style type='text/css'>" + css + "</style></defs>" + content + "</svg>");
}
