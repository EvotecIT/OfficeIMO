using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsStyleTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    public void BoldAddsTwoPercentEmOnlyToImplicitAdvances(bool sideways, bool rtl) {
        var doc = Create("None", sideways, rtl); var page = doc.Pages[0];
        double[] original = Points(page); int count = original.Length / 2;
        var xml = page.GetMarkup(); var glyph = xml.Elements().Single();
        glyph.SetAttributeValue("StyleSimulations", "BoldSimulation"); page.ReplaceMarkup(xml);
        var bold = Points(page);
        Assert.Equal(original.Length, bold.Length);
        Assert.Equal(original[count] - original[0] + (rtl ? -2 : 2), bold[count] - bold[0], 3);
        glyph.SetAttributeValue("Indices", ",70;,70"); page.ReplaceMarkup(xml); bold = Points(page);
        Assert.Equal(rtl ? -70 : 70, bold[count] - bold[0], 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ItalicSkewsEachGlyphAboutItsNativeOrigin(bool sideways) {
        var doc = Create("None", sideways); var page = doc.Pages[0];
        var xml = page.GetMarkup(); var glyph = xml.Elements().Single(); glyph.SetAttributeValue("UnicodeString", "I");
        page.ReplaceMarkup(xml); var original = Points(page);
        glyph.SetAttributeValue("StyleSimulations", "ItalicSimulation"); page.ReplaceMarkup(xml); var italic = Points(page);
        double shear = Math.Tan(20 * Math.PI / 180);
        var tables = XpsSidewaysFixtures.Tables();
        double topY = XpsSidewaysFixtures.I16(tables["OS/2"], 68) * 100D / XpsSidewaysFixtures.U16(tables["head"], 18);
        for (int i = 0; i < original.Length; i += 2) {
            Assert.Equal(sideways ? original[i] : original[i] - shear * (original[i + 1] - 120), italic[i], 3);
            Assert.Equal(sideways ? original[i + 1] + shear * (original[i] - 30 - topY) : original[i + 1], italic[i + 1], 3);
        }
    }

    [Theory]
    [InlineData("BoldSimulation", XpsFormat.Xps)]
    [InlineData("BoldItalicSimulation", XpsFormat.OpenXps)]
    public void OverlappingSimulatedGlyphsPaintTranslucentBrushOnce(string simulation, XpsFormat format) {
        var doc = Create(simulation, format: format); var page = doc.Pages[0];
        var xml = page.GetMarkup(); var glyph = xml.Elements().Single(); glyph.SetAttributeValue("Indices", ",0;,0"); glyph.SetAttributeValue("Fill", "#80FF0000");
        if (simulation == "BoldItalicSimulation") {
            glyph.Attribute("Fill")!.Remove(); var ns = xml.Name.Namespace;
            glyph.Add(new XElement(ns + "Glyphs.Fill", new XElement(ns + "LinearGradientBrush", new XAttribute("MappingMode", "Absolute"), new XAttribute("StartPoint", "0,0"), new XAttribute("EndPoint", "300,0"),
                new XElement(ns + "LinearGradientBrush.GradientStops", new XElement(ns + "GradientStop", new XAttribute("Offset", "0"), new XAttribute("Color", "#80FF0000")),
                    new XElement(ns + "GradientStop", new XAttribute("Offset", "1"), new XAttribute("Color", "#80FF0000"))))));
        }
        page.ReplaceMarkup(xml); byte[] bytes = XpsDocument.Load(doc.Save()).Pages[0].ExportImage(OfficeImageExportFormat.Png).Bytes;
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out var raster));
        byte minimum = 255;
        for (int y = 0; y < raster!.Height; y++) for (int x = 0; x < raster.Width; x++) minimum = Math.Min(minimum, raster.GetPixel(x, y).G);
        Assert.InRange(minimum, (byte)126, (byte)129);
        Assert.NotEmpty(doc.ToPdf());
    }

    [Fact]
    public void InvalidSimulationIsNotTreatedAsAnUnstyledRun() {
        Assert.Throws<InvalidDataException>(() => Create("Unknown").Pages[0].ToSvg(true));
    }

    [Fact]
    public void GlyphOpacityMaskComposesWithBoldCoverage() {
        var page = Create("BoldSimulation").Pages[0]; var xml = page.GetMarkup();
        xml.Elements().Single().SetAttributeValue("OpacityMask", "#80000000"); page.ReplaceMarkup(xml);
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ExportImage(OfficeImageExportFormat.Png).Bytes, out var raster));
        Assert.Equal(OfficeColor.White, raster!.GetPixel(0, 0));
        byte darkest = 255;
        for (int y = 0; y < raster.Height; y++) for (int x = 0; x < raster.Width; x++) darkest = Math.Min(darkest, raster.GetPixel(x, y).R);
        Assert.InRange(darkest, (byte)126, (byte)129);
    }

    [Fact]
    public void DocumentPageWithManySmallBoldRunsUsesLocalCoverageSurfaces() {
        var doc = XpsDocument.Create(); string font = doc.AddFont(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "RobotoFlex.ttf")));
        var page = doc.AddPage(816, 1056);
        for (int i = 0; i < 30; i++) page.AddText("Heading", font, 14, 30, 30 + i * 25);
        var xml = page.GetMarkup(); foreach (var glyph in xml.Elements()) glyph.SetAttributeValue("StyleSimulations", "BoldSimulation"); page.ReplaceMarkup(xml);
        Assert.NotNull(page.ToDrawing()); Assert.NotEmpty(doc.ToPdf());
    }

    private static XpsDocument Create(string simulation, bool sideways = false, bool rtl = false, XpsFormat format = XpsFormat.OpenXps) {
        var doc = XpsDocument.Create(format); string font = doc.AddFont(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "RobotoFlex.ttf")));
        var page = doc.AddPage(300, 180).AddText("II", font, 100, rtl ? 250 : 30, 120);
        var xml = page.GetMarkup(); var run = xml.Elements().Single(); run.SetAttributeValue("StyleSimulations", simulation);
        run.SetAttributeValue("IsSideways", sideways ? "true" : "false"); run.SetAttributeValue("BidiLevel", rtl ? "1" : "0"); page.ReplaceMarkup(xml);
        return doc;
    }
    private static double[] Points(XpsPage page) {
        var svg = XElement.Parse(page.ToSvg().Svg);
        string data = (string)svg.Descendants().Single(e => e.Name.LocalName == "path" && (string?)e.Attribute("fill-rule") == "nonzero").Attribute("d")!;
        return Regex.Matches(data, @"-?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?").Cast<Match>().Select(m => double.Parse(m.Value, CultureInfo.InvariantCulture)).ToArray();
    }
}
