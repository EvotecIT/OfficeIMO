using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsDocumentTests {
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static byte[] Font => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "RobotoFlex.ttf"));
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void CreatesReopensEditsAndPreservesNativeParts(XpsFormat format) {
        var document = XpsDocument.Create(format);
        string font = document.AddFont(Font);
        document.AddResource("Metadata/opaque.bin", new byte[] { 3, 7, 9 }, "application/octet-stream");
        document.AddPage(320, 160).AddPath("M10,10 L310,10 310,150 10,150Z", "#FF0080FF").AddText("{Hello XPS}", font, 24, 30, 80);
        Assert.Equal(document.Save(), document.Save());
        using var stream = new MemoryStream(document.Save());
        var loaded = XpsDocument.Load(stream);
        Assert.True(stream.CanRead);
        Assert.Equal(format, loaded.Format);
        Assert.Equal("{Hello XPS}", loaded.Pages[0].ExtractText());
        var markup = loaded.Pages[0].GetMarkup();
        markup.Elements().First().SetAttributeValue("Fill", "#FF00FF00");
        loaded.Pages[0].ReplaceMarkup(markup);
        var reopened = XpsDocument.Load(loaded.Save());
        Assert.Equal(new byte[] { 3, 7, 9 }, reopened.GetPartBytes("Metadata/opaque.bin"));
        Assert.Equal("#FF00FF00", (string?)reopened.Pages[0].GetMarkup().Elements().First().Attribute("Fill"));
        Assert.True(reopened.Pages[0].ToSvg().IsComplete);
        Assert.DoesNotContain("<text", reopened.Pages[0].ToSvg().Svg);
        byte[] png = reopened.Pages[0].ExportImage(OfficeImageExportFormat.Png).Bytes;
        Assert.Equal(320, OfficeImageReader.Identify(png).Width);
        byte[] pdf = reopened.ToPdf(); Assert.StartsWith("%PDF-", Encoding.ASCII.GetString(pdf, 0, 8));
    }
    [Fact]
    public void RejectsExternalStartRelationshipAndMissingPages() {
        byte[] bytes = Sample();
        Assert.Throws<InvalidDataException>(() => XpsDocument.Load(Rewrite(bytes, "_rels/.rels", b => {
            var xml = XElement.Parse(Encoding.UTF8.GetString(b)); xml.Elements().Single().SetAttributeValue("TargetMode", "External"); return Encoding.UTF8.GetBytes(xml.ToString());
        })));
        Assert.Throws<InvalidDataException>(() => XpsDocument.Load(Rewrite(bytes, "Documents/1/FixedDocument.fdoc", b => Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(b).Replace("1.fpage", "missing.fpage")))));
    }
    [Fact]
    public void BoundsInputExpansionPagesAndXmlDepth() {
        var doc = XpsDocument.Create(); doc.AddPage(); doc.AddPage(); byte[] bytes = doc.Save();
        Assert.Throws<InvalidDataException>(() => XpsDocument.Load(bytes, new XpsReadOptions { MaximumInputBytes = 20 }));
        Assert.Throws<InvalidDataException>(() => XpsDocument.Load(bytes, new XpsReadOptions { MaximumExpandedBytes = 50 }));
        Assert.Throws<InvalidDataException>(() => XpsDocument.Load(bytes, new XpsReadOptions { MaximumPages = 1 }));
        Assert.Throws<InvalidDataException>(() => XpsDocument.Load(bytes, new XpsReadOptions { MaximumParts = 1 }));
        Assert.Throws<InvalidDataException>(() => XpsDocument.Load(Rewrite(bytes, "Documents/1/Pages/1.fpage", b => Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(b).Replace(" />", "><Canvas><Canvas /></Canvas></FixedPage>"))), new XpsReadOptions { MaximumXmlDepth = 1 }));
        Assert.Throws<OperationCanceledException>(() => XpsDocument.Load(bytes, cancellationToken: new CancellationToken(true)));
    }
    [Fact]
    public void RejectsDtdDuplicatePartsAndTraversal() {
        Assert.ThrowsAny<System.Xml.XmlException>(() => XpsDocument.Load(Rewrite(Sample(), "_rels/.rels", b => Encoding.UTF8.GetBytes("<!DOCTYPE x [<!ENTITY a 'a'>]>" + Encoding.UTF8.GetString(b).Substring(Encoding.UTF8.GetString(b).IndexOf("<Relationships", StringComparison.Ordinal))))));
        Assert.Throws<InvalidDataException>(() => XpsDocument.Load(WithExtra(Sample(), "../escape", new byte[] { 1 })));
        Assert.Throws<InvalidDataException>(() => XpsDocument.Load(WithExtra(Sample(), "Documents/1/Pages/1.fpage", new byte[] { 1 })));
        Assert.Throws<InvalidDataException>(() => XpsDocument.Load(WithExtra(Sample(), "%2e%2e/escape", new byte[] { 1 })));
    }
    [Fact]
    public void NativeUnsupportedMarkupIsPreservedButStrictConversionRefusesIt() {
        var doc = XpsDocument.Load(Sample()); var page = doc.Pages[0]; var xml = page.GetMarkup();
        xml.Elements().First().SetAttributeValue("Stroke", "#FF000000");
        xml.Elements().First().SetAttributeValue("StrokeStartLineCap", "Triangle"); page.ReplaceMarkup(xml);
        Assert.True(page.ToSvg().IsComplete);
        Assert.NotNull(XpsDocument.Load(doc.Save()).Pages[0].GetMarkup().Elements().First().Attribute("StrokeStartLineCap"));
    }
    [Fact]
    public void SignedPackagesCannotBeRewrittenAndAtomicDestinationSurvives() {
        var doc = XpsDocument.Create(); doc.AddPage(); doc.AddResource("_xmlsignatures/sig.xml", Encoding.UTF8.GetBytes("<sig/>"), "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml");
        string path = Path.GetTempFileName();
        try { File.WriteAllText(path, "original"); Assert.Throws<NotSupportedException>(() => doc.Save(path)); Assert.Equal("original", File.ReadAllText(path)); }
        finally { File.Delete(path); }
    }
    [Fact]
    public void ResolvesResourcesTransformsClipsGradientsAndLinks() {
        var doc = XpsDocument.Create(); var page = doc.AddPage(100, 100); var xml = page.GetMarkup(); var ns = xml.Name.Namespace;
        xml.Add(new XElement(ns + "FixedPage.Resources", new XElement(ns + "ResourceDictionary", new XElement(ns + "SolidColorBrush", new XAttribute(XName.Get("Key", "http://schemas.openxps.org/oxps/v1.0/resourcedictionary-key"), "blue"), new XAttribute("Color", "#800000FF")))));
        xml.Add(new XElement(ns + "Canvas", new XAttribute("RenderTransform", "1,0,0,1,10,20"), new XAttribute("Clip", "M0,0L50,0 50,50Z"),
            new XElement(ns + "Path", new XAttribute("Data", "M0,0L40,0 40,40Z"), new XAttribute("Fill", "{StaticResource blue}"), new XAttribute("FixedPage.NavigateUri", "https://example.com"))));
        page.ReplaceMarkup(xml); var svg = XElement.Parse(page.ToSvg().Svg); XNamespace sns = "http://www.w3.org/2000/svg";
        Assert.Equal("https://example.com", (string?)svg.Descendants(sns + "a").Single().Attribute("href"));
        Assert.Contains(svg.Descendants(), e => (string?)e.Attribute("transform") == "matrix(1 0 0 1 10 20)");
        Assert.Contains(svg.Descendants(sns + "path"), e => (string?)e.Attribute("fill") == "#0000FF");
    }
    [Fact]
    public void GlyphIndicesOverrideUnicodeAndAdvancesRemainPositioned() {
        var doc = XpsDocument.Create(); string font = doc.AddFont(Font, false); var page = doc.AddPage(200, 100).AddText("AB", font, 24, 10, 50);
        var xml = page.GetMarkup(); var glyphs = xml.Elements().Single(); glyphs.SetAttributeValue("Indices", "36,100;37,100"); page.ReplaceMarkup(xml);
        string first = page.ToSvg().Svg;
        glyphs.SetAttributeValue("UnicodeString", "ZZ"); page.ReplaceMarkup(xml);
        string second = page.ToSvg().Svg;
        Assert.Equal(first.Replace("aria-label=\"AB\"", "aria-label=\"ZZ\""), second);
        glyphs.SetAttributeValue("Indices", "65535;37"); page.ReplaceMarkup(xml);
        Assert.Throws<ArgumentOutOfRangeException>(() => page.ToSvg());
    }
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void ImagesRetainPlacementAndPixels(XpsFormat dialect) {
        var doc = XpsDocument.Create(dialect);
        string image = doc.AddResource("Resources/test.png", OfficePngWriter.Encode(new OfficeRasterImage(4, 2, OfficeColor.SteelBlue)), "image/png");
        var page = doc.AddPage(80, 60).AddImage(image, 10, 20, 40, 20);
        var loaded = XpsDocument.Load(doc.Save());
        Assert.True(OfficeRasterImageDecoder.TryDecode(loaded.Pages[0].ExportImage(OfficeImageExportFormat.Png).Bytes, out var raster));
        Assert.NotNull(raster);
        Assert.Equal(OfficeColor.SteelBlue, raster!.GetPixel(30, 30));
        Assert.NotEqual(OfficeColor.SteelBlue, raster.GetPixel(5, 5));
    }
    [Fact]
    public void ExternalResourceDictionaryUsesItsOwnBaseUri() {
        var doc = XpsDocument.Create(); var page = doc.AddPage(80, 60); var ns = page.GetMarkup().Name.Namespace;
        string keyNs = "http://schemas.openxps.org/oxps/v1.0/resourcedictionary-key";
        doc.AddResource("Resources/Images/test.png", OfficePngWriter.Encode(new OfficeRasterImage(4, 2, OfficeColor.SteelBlue)), "image/png");
        var dictionary = new XElement(ns + "ResourceDictionary", new XElement(ns + "ImageBrush", new XAttribute(XName.Get("Key", keyNs), "image"),
            new XAttribute("ImageSource", "../Images/test.png"), new XAttribute("Viewbox", "0,0,4,2"), new XAttribute("Viewport", "10,20,40,20"), new XAttribute("ViewboxUnits", "Absolute"), new XAttribute("ViewportUnits", "Absolute")));
        doc.AddResource("Resources/Dictionaries/brush.xaml", Encoding.UTF8.GetBytes(dictionary.ToString()), "application/vnd.ms-package.xps-resourcedictionary+xml");
        var markup = page.GetMarkup();
        markup.Add(new XElement(ns + "FixedPage.Resources", new XElement(ns + "ResourceDictionary", new XAttribute("Source", "/Resources/Dictionaries/brush.xaml"))),
            new XElement(ns + "Path", new XAttribute("Data", "M0,0H80V60H0Z"), new XAttribute("Fill", "{StaticResource image}")));
        page.ReplaceMarkup(markup);
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ExportImage(OfficeImageExportFormat.Png).Bytes, out var raster));
        Assert.Equal(OfficeColor.SteelBlue, raster!.GetPixel(30, 30));
    }
    [Fact]
    public void RejectsResourceCyclesAndInvalidGeometry() {
        var doc = XpsDocument.Create(); var page = doc.AddPage(); var xml = page.GetMarkup(); XNamespace ns = xml.Name.Namespace;
        doc.AddResource("Resources/cycle.xaml", Encoding.UTF8.GetBytes(new XElement(ns + "ResourceDictionary", new XAttribute("Source", "cycle.xaml")).ToString()), "application/vnd.ms-package.xps-resourcedictionary+xml");
        xml.Add(new XElement(ns + "FixedPage.Resources", new XElement(ns + "ResourceDictionary", new XAttribute("Source", "/Resources/cycle.xaml")))); page.ReplaceMarkup(xml);
        Assert.Throws<InvalidDataException>(() => page.ToSvg());
        var invalid = XpsDocument.Create().AddPage().AddPath("M0,0 C1,2");
        Assert.Throws<InvalidDataException>(() => invalid.ToSvg());
    }
    [Fact]
    public void ClustersCanMapSeveralCodeUnitsToOneExplicitGlyph() {
        var doc = XpsDocument.Create(); var page = doc.AddPage(150, 80).AddText("fi", doc.AddFont(Font), 24, 10, 40);
        var xml = page.GetMarkup(); xml.Elements().Single().SetAttributeValue("Indices", "(2:1)36,100"); page.ReplaceMarkup(xml);
        Assert.True(page.ToSvg().IsComplete);
        xml.Elements().Single().SetAttributeValue("Indices", "(3:1)36"); page.ReplaceMarkup(xml);
        Assert.Throws<InvalidDataException>(() => page.ToSvg());
    }
    [Fact]
    public void GradientStopsAndOpacityUseNativeCoordinates() {
        var doc = XpsDocument.Create(); var page = doc.AddPage(100, 50); var xml = page.GetMarkup(); XNamespace ns = xml.Name.Namespace;
        xml.Add(new XElement(ns + "Path", new XAttribute("Data", "M0,0H100V50H0Z"), new XElement(ns + "Path.Fill",
            new XElement(ns + "LinearGradientBrush", new XAttribute("StartPoint", "0,0"), new XAttribute("EndPoint", "100,0"), new XAttribute("MappingMode", "Absolute"),
                new XElement(ns + "LinearGradientBrush.GradientStops", new XElement(ns + "GradientStop", new XAttribute("Offset", "0"), new XAttribute("Color", "#FFFF0000")), new XElement(ns + "GradientStop", new XAttribute("Offset", "1"), new XAttribute("Color", "#FF0000FF")))))));
        page.ReplaceMarkup(xml);
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ExportImage(OfficeImageExportFormat.Png).Bytes, out var raster));
        Assert.True(raster!.GetPixel(10, 20).R > raster.GetPixel(90, 20).R);
        Assert.True(raster.GetPixel(10, 20).B < raster.GetPixel(90, 20).B);
    }
    [Theory]
    [InlineData("F1")]
    [InlineData("F 1")]
    [InlineData("F\t1")]
    public void PathFillRuleAllowsSpecificationWhitespace(string prefix) {
        var page = XpsDocument.Create().AddPage(40, 40).AddPath(prefix + " M0,0 H30 V30 H0Z");
        Assert.Contains("fill-rule=\"nonzero\"", page.ToSvg().Svg);
    }
    [Fact]
    public void WhitespaceOnlyGlyphRunsDoNotBecomeFailedDrawingGeometry() {
        var doc = XpsDocument.Create(); var page = doc.AddPage(100, 100).AddText(" ", doc.AddFont(Font), 16, 10, 30);
        Assert.Equal(" ", page.ExtractText());
        Assert.NotNull(page.ToDrawing());
    }
    private static byte[] Sample() { var d = XpsDocument.Create(); d.AddPage(100, 100).AddPath("M0,0L50,0 50,50Z"); return d.Save(); }
    internal static byte[] Rewrite(byte[] bytes, string name, Func<byte[], byte[]> change) {
        using var output = new MemoryStream(); using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true)) {
            using var input = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
            foreach (var part in input.Entries) { using var source = part.Open(); using var buffer = new MemoryStream(); source.CopyTo(buffer); byte[] data = buffer.ToArray(); if (part.FullName == name) data = change(data); using var target = zip.CreateEntry(part.FullName).Open(); target.Write(data, 0, data.Length); }
        } return output.ToArray();
    }
    private static byte[] WithExtra(byte[] bytes, string name, byte[] data) {
        using var stream = new MemoryStream(); stream.Write(bytes, 0, bytes.Length); using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, true)) { using var part = zip.CreateEntry(name).Open(); part.Write(data, 0, data.Length); } return stream.ToArray();
    }
}
