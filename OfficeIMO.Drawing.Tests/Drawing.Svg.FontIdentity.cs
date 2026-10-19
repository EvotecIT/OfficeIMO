using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingSvgFontImportTests {
    [Theory]
    [InlineData("SVG Fixture, \"Sans\" & Co")]
    [InlineData("Fixture\\Slash")]
    public void ScopedSingleFamilyIdentitySurvivesFallbackMeasurementAndPainting(string family) {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Carlito-Regular.ttf"));
        OfficeDrawing Make(string name) {
            var drawing = new OfficeDrawing(180, 50);
            drawing.Fonts.Add(name, font);
            drawing.AddText("Wiii 0123", 8, 5, 160, 40, new OfficeFontInfo(OfficeFontFamilyParser.Quote(name), 24));
            return drawing;
        }
        OfficeDrawing plain = Make("PlainFixture");
        OfficeDrawing aliased = Make(family);
        Assert.True(plain.Fonts.TryMeasureText("Wiii 0123", 24, "PlainFixture", OfficeFontStyle.Regular, out double expectedWidth));
        Assert.True(aliased.Fonts.TryMeasureText("Wiii 0123", 24, OfficeFontFamilyParser.Quote(family), OfficeFontStyle.Regular, out double actualWidth));
        Assert.Equal(expectedWidth, actualWidth, 6);
        Assert.Equal(OfficeDrawingRasterRenderer.Render(plain).GetPixels(), OfficeDrawingRasterRenderer.Render(aliased).GetPixels());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PunctuatedScopedFamilyRetainsEligibleSvgOutlineExport(bool painted) {
        const string family = "SVG Fixture, \"Sans\" & Co";
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add(family, OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.AddText("A", 4, 4, 90, 30, new OfficeFontInfo(OfficeFontFamilyParser.Quote(family), 14));
        if (painted) drawing.MarkLastTextAsPaintedGlyphs();
        else drawing.SvgTextAsPaths = true;

        var svg = System.Xml.Linq.XElement.Parse(OfficeDrawingSvgExporter.ToSvg(drawing));
        Assert.Single(svg.Descendants().Where(element => element.Name.LocalName == "path"));
        Assert.DoesNotContain(svg.Descendants(), element => element.Name.LocalName == "text");
    }

    [Fact]
    public void SupplementaryCssEscapeNamesMatchDeclarationsAndTextReferences() {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Carlito-Regular.ttf"));
        OfficeDrawing Read(string family) {
            string css = FontFace(family, font);
            byte[] svg = SvgFontImportDocument(css, "<text x='8' y='28' font-family='\"" + family + "\"'>Wiii</text>");
            Assert.True(OfficeSvgDrawingReader.TryRead(svg, out OfficeDrawing? drawing, out int unsupported));
            Assert.Equal(0, unsupported);
            return drawing!;
        }
        OfficeDrawing actual = Read("Fixture\\1F600 ");
        OfficeDrawing expected = Read("PlainFixture");
        Assert.Contains(actual.Fonts.Faces, face => face.FamilyName == "Fixture\U0001F600");
        Assert.Equal(OfficeDrawingRasterRenderer.Render(expected).GetPixels(), OfficeDrawingRasterRenderer.Render(actual).GetPixels());
    }
}
