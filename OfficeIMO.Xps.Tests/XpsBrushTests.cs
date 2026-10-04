using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsBrushTests {
    [Theory]
    [InlineData("image-Tile", false)]
    [InlineData("image-FlipX", true)]
    [InlineData("image-FlipY", false)]
    [InlineData("image-FlipXY", true)]
    [InlineData("visual-Tile", false)]
    [InlineData("visual-FlipX", true)]
    [InlineData("visual-FlipY", false)]
    [InlineData("visual-FlipXY", true)]
    public void TilesRepeatWithNativeFlipParity(string scenario, bool flipped) {
        var page = XpsBrushFixtures.Create(scenario).Pages[0];
        var raster = Raster(page);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(15, 15));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(25, 15));
        Assert.Equal(flipped ? OfficeColor.Blue : OfficeColor.Red, raster.GetPixel(35, 15));
        Assert.Equal(flipped ? OfficeColor.Red : OfficeColor.Blue, raster.GetPixel(45, 15));
        Assert.Equal(OfficeColor.Green, raster.GetPixel(15, 25));
        bool flipY = scenario.EndsWith("FlipY", StringComparison.Ordinal) || scenario.EndsWith("FlipXY", StringComparison.Ordinal);
        Assert.Equal(flipY ? OfficeColor.Green : OfficeColor.Red, raster.GetPixel(15, 35));
    }
    [Theory]
    [InlineData("image-None")]
    [InlineData("visual-None")]
    public void NonTiledBrushOnlyPaintsItsViewport(string scenario) {
        var raster = Raster(XpsBrushFixtures.Create(scenario).Pages[0]);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(15, 15));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(25, 15));
        Assert.NotEqual(OfficeColor.Red, raster.GetPixel(35, 15));
    }
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void GradientTransformSurvivesPackageAndPdfConversion(XpsFormat format) {
        var doc = XpsDocument.Load(XpsBrushFixtures.Create("gradient", format).Save());
        var raster = Raster(doc.Pages[0]);
        Assert.True(raster.GetPixel(21, 40).R > raster.GetPixel(59, 40).R);
        Assert.True(raster.GetPixel(59, 40).B > raster.GetPixel(21, 40).B);
        Assert.NotEmpty(doc.ToPdf());
    }
    [Theory]
    [InlineData("mask")]
    [InlineData("mask-transformed")]
    public void OpacityMasksUseAlphaAndFollowTheLocalCoordinateSpace(string scenario) {
        var raster = Raster(XpsBrushFixtures.Create(scenario).Pages[0]);
        // Black is opaque in an alpha mask; a luminosity mask would erase both samples.
        Assert.True(raster.GetPixel(90, 40).R > 200);
        Assert.True(raster.GetPixel(10, 40).G > raster.GetPixel(90, 40).G || raster.GetPixel(10, 40).A < raster.GetPixel(90, 40).A);
    }
    [Theory]
    [InlineData("image-mask")]
    [InlineData("visual-mask")]
    public void TiledBrushMasksUseBrushAlpha(string scenario) {
        var raster = Raster(XpsBrushFixtures.Create(scenario).Pages[0]);
        foreach (int y in new[] { 15, 25 }) {
            var pixel = raster.GetPixel(15, y);
            Assert.True((pixel.A >= 126 && pixel.A <= 129) || (pixel.G >= 126 && pixel.G <= 129));
        }
    }
    [Fact]
    public void TileTransformMovesThePatternOrigin() {
        var raster = Raster(XpsBrushFixtures.Create("visual-transformed").Pages[0]);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(20, 20));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(30, 20));
    }
    [Fact]
    public void VisualBrushReferencesAndCyclesRemainBounded() {
        var doc = XpsBrushFixtures.Create("visual-Tile", XpsFormat.OpenXps); var page = doc.Pages[0]; var xml = page.GetMarkup(); var ns = xml.Name.Namespace;
        XNamespace key = "http://schemas.openxps.org/oxps/v1.0/resourcedictionary-key";
        var brush = xml.Descendants(ns + "VisualBrush").Single(); var visual = brush.Element(ns + "VisualBrush.Visual")!.Elements().Single(); visual.Remove(); visual.SetAttributeValue(key + "Key", "visual");
        brush.Element(ns + "VisualBrush.Visual")!.Remove(); brush.SetAttributeValue("Visual", "{StaticResource visual}");
        xml.AddFirst(new XElement(ns + "FixedPage.Resources", new XElement(ns + "ResourceDictionary", visual))); page.ReplaceMarkup(xml);
        Assert.Equal(OfficeColor.Red, Raster(page).GetPixel(15, 15));
        // A resource visual referencing the same brush must terminate before expansion.
        brush.Remove(); brush.SetAttributeValue(key + "Key", "brush"); xml.Elements().First().Elements().Single().Add(brush);
        visual.Elements().First().SetAttributeValue("Fill", "{StaticResource brush}");
        xml.Elements().Last().SetAttributeValue("Fill", "{StaticResource brush}"); xml.Elements().Last().Elements().Remove(); page.ReplaceMarkup(xml);
        Assert.Throws<InvalidDataException>(() => page.ToSvg());
    }
    [Fact]
    public void CanvasMaskAppliesOnceToOverlappingChildren() {
        var doc = XpsDocument.Create(); var page = doc.AddPage(80, 60); var xml = page.GetMarkup(); var ns = xml.Name.Namespace;
        XElement Path() => new(ns + "Path", new XAttribute("Data", "M10,10H60V50H10Z"), new XAttribute("Fill", "#FFFF0000"));
        xml.Add(new XElement(ns + "Canvas", new XAttribute("OpacityMask", "#80000000"), Path(), Path())); page.ReplaceMarkup(xml);
        var pixel = Raster(page).GetPixel(30, 30);
        Assert.True((pixel.A >= 126 && pixel.A <= 129) || (pixel.G >= 126 && pixel.G <= 129));
    }

    [Fact]
    public void GroupOpacityAndMaskComposeAfterOverlappingChildren() {
        var pixel = Raster(XpsBrushFixtures.Create("mask-group-opacity").Pages[0]).GetPixel(30, 30);
        Assert.True((pixel.A >= 63 && pixel.A <= 65) || (pixel.G >= 190 && pixel.G <= 193));
    }

    [Fact]
    public void ZeroAreaBrushPaintsNothingAndVisualExpansionIsBounded() {
        var page = XpsBrushFixtures.Create("visual-Tile").Pages[0]; var xml = page.GetMarkup(); var ns = xml.Name.Namespace;
        var brush = xml.Descendants(ns + "VisualBrush").Single(); brush.SetAttributeValue("Viewport", "0,0,0,10"); page.ReplaceMarkup(xml);
        Assert.NotEqual(OfficeColor.Red, Raster(page).GetPixel(15, 15));
        brush.SetAttributeValue("Viewport", "0,0,10,10");
        var path = xml.Elements().Single();
        // Repeated vector tiles cannot retain unlimited expanded SVG nodes.
        var visual = brush.Element(ns + "VisualBrush.Visual")!.Elements().Single();
        for (int i = 0; i < 400; i++) visual.Add(new XElement(ns + "Path", new XAttribute("Data", "M0,0L1,1"), new XAttribute("Fill", "#FF000000")));
        for (int i = 0; i < 260; i++) xml.Add(new XElement(path));
        page.ReplaceMarkup(xml);
        Assert.Throws<InvalidDataException>(() => page.ToSvg(true));
    }

    private static OfficeRasterImage Raster(XpsPage page) {
        Assert.True(page.ToSvg().IsComplete);
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ExportImage(OfficeImageExportFormat.Png).Bytes, out var raster));
        return raster!;
    }
}
