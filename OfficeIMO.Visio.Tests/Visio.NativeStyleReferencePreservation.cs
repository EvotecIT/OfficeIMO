using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioNativeStyleReferencePreservationTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";

    [Fact]
    public void IndependentDrawingRetainsShapeAndDocumentStyleReferences() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", "shorewall-netfilter.vdx");
        VisioDocument document = VisioDocument.LoadLegacyXml(path).Value;
        foreach (VisioDocument candidate in new[] { document, ReopenPackage(document), ReopenLegacy(document) }) {
            XDocument xml = Export(candidate);
            XElement shape = xml.Descendants(Legacy + "Page").Descendants(Legacy + "Shape")
                .Single(element => (string?)element.Attribute("ID") == "32");
            AssertReferences(shape, "1", "1", "6");
            Assert.Equal("6", (string?)xml.Root!.Element(Legacy + "DocumentSettings")!.Attribute("DefaultTextStyle"));
        }
    }

    [Fact]
    public void LoadedShapesAndConnectorsRetainExplicitPartialAndAbsentReferences() {
        VisioDocument document = Load(StyleHeaders + Pages(
            Shape("1", "LineStyle='001' FillStyle='9' TextStyle='06'")
            + Shape("2", "TextStyle='6'") + Shape("3", "")
            + Connector("4", "LineStyle='7' TextStyle='6'") + Connector("5", "")));
        foreach (VisioDocument candidate in new[] { document, ReopenPackage(document), ReopenLegacy(document) }) {
            foreach (XDocument xml in new[] { Export(candidate), ReadPart(candidate, "visio/pages/page1.xml") }) {
                XNamespace ns = xml.Root!.Name.Namespace;
                XElement[] shapes = xml.Descendants(ns + "Shape").ToArray();
                AssertReferences(shapes.Single(shape => (string?)shape.Attribute("ID") == "1"), "001", "9", "06");
                AssertReferences(shapes.Single(shape => (string?)shape.Attribute("ID") == "2"), null, null, "6");
                AssertReferences(shapes.Single(shape => (string?)shape.Attribute("ID") == "3"), null, null, null);
                AssertReferences(shapes.Single(shape => (string?)shape.Attribute("ID") == "4"), "7", null, "6");
                AssertReferences(shapes.Single(shape => (string?)shape.Attribute("ID") == "5"), null, null, null);
            }
        }
    }

    [Fact]
    public void NativeGeneratedStyleHeadersAndDocumentDefaultsOverrideAuthoredDefaults() {
        VisioDocument document = Load("<DocumentSettings DefaultTextStyle='6' DefaultLineStyle='7' DefaultFillStyle='8' DefaultGuideStyle='9'/>"
            + "<StyleSheets><StyleSheet ID='0' TextStyle='6'><StyleProp><EnableTextProps>0</EnableTextProps></StyleProp></StyleSheet>"
            + "<StyleSheet ID='1' LineStyle='7' FillStyle='8' TextStyle='6'/><StyleSheet ID='2'/>"
            + "<StyleSheet ID='6'/><StyleSheet ID='7'/><StyleSheet ID='8'/><StyleSheet ID='9'/></StyleSheets>"
            + Pages(Shape("1", "")));
        byte[] native = document.ToBytes();
        using (var stream = new MemoryStream()) {
            stream.Write(native, 0, native.Length);
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, true)) {
                ZipArchiveEntry entry = archive.GetEntry("visio/document.xml")!;
                XDocument xml;
                using (var input = entry.Open()) xml = XDocument.Load(input);
                xml.Descendants(Modern + "StyleSheet").Single(style => (string?)style.Attribute("ID") == "1")
                    .SetAttributeValue("BasedOn", "9");
                entry.Delete();
                using var output = archive.CreateEntry("visio/document.xml").Open();
                xml.Save(output);
            }
            native = stream.ToArray();
        }
        document = VisioDocument.Load(new MemoryStream(native));
        foreach (VisioDocument candidate in new[] { document, ReopenPackage(document) }) {
            XDocument xml = ReadPart(candidate, "visio/document.xml");
            XElement settings = xml.Root!.Element(Modern + "DocumentSettings")!;
            Assert.Equal("6", (string?)settings.Attribute("DefaultTextStyle"));
            Assert.Equal("7", (string?)settings.Attribute("DefaultLineStyle"));
            Assert.Equal("8", (string?)settings.Attribute("DefaultFillStyle"));
            Assert.Equal("9", (string?)settings.Attribute("DefaultGuideStyle"));
            XElement[] styles = xml.Descendants(Modern + "StyleSheet").ToArray();
            XElement zero = styles.Single(style => (string?)style.Attribute("ID") == "0");
            AssertReferences(zero, null, null, "6");
            Assert.Null(zero.Attribute("BasedOn"));
            Assert.Equal("0", (string?)Assert.Single(zero.Elements(Modern + "Cell"),
                cell => (string?)cell.Attribute("N") == "EnableTextProps").Attribute("V"));
            XElement one = styles.Single(style => (string?)style.Attribute("ID") == "1");
            AssertReferences(one, "7", "8", "6");
            Assert.Equal("9", (string?)one.Attribute("BasedOn"));
            XElement two = styles.Single(style => (string?)style.Attribute("ID") == "2");
            AssertReferences(two, null, null, null);
            Assert.Null(two.Attribute("BasedOn"));
        }
    }

    [Fact]
    public void PageDuplicationAndMasterInstancesKeepReferencesAndDynamicMasterInheritance() {
        VisioDocument document = Load(StyleHeaders
            + "<Masters><Master ID='8' NameU='Native box'><Shapes>"
            + Shape("1", "LineStyle='7' FillStyle='8' TextStyle='6'") + "</Shapes></Master></Masters>"
            + Pages(Shape("1", "LineStyle='7' FillStyle='8' TextStyle='6'") + Shape("2", "")
                + Connector("3", "NameU='Dynamic connector'") + Connector("4", "TextStyle='6'")));
        VisioPage source = document.Pages[0];
        source.AddShape("5", document.GetMaster("Native box"), 4, 4, 2, 1);
        document.DuplicatePage(source, "Copy");
        foreach (VisioDocument candidate in new[] { document, ReopenPackage(document), ReopenLegacy(document) }) {
            XDocument xml = Export(candidate);
            Assert.Equal(2, xml.Descendants(Legacy + "Page").Count());
            foreach (XElement page in xml.Descendants(Legacy + "Page")) {
                XElement[] shapes = page.Element(Legacy + "Shapes")!.Elements(Legacy + "Shape").ToArray();
                AssertReferences(shapes.Single(shape => (string?)shape.Attribute("NameU") == "Node1"), "7", "8", "6");
                AssertReferences(shapes.Single(shape => (string?)shape.Attribute("NameU") == "Node2"), null, null, null);
                XElement dynamic = shapes.Single(shape => (string?)shape.Attribute("NameU") == "Dynamic connector");
                AssertReferences(dynamic, null, null, null);
                Assert.NotNull(dynamic.Attribute("Master"));
                AssertReferences(shapes.Single(shape => shape.Element(Legacy + "Text")?.Value == "Label4"), null, null, "6");
                AssertReferences(shapes.Single(shape => (string?)shape.Attribute("NameU") == "Native box"), "7", "8", "6");
            }
        }
    }

    private const string StyleHeaders = "<StyleSheets><StyleSheet ID='6'/><StyleSheet ID='7'/><StyleSheet ID='8'/><StyleSheet ID='9'/></StyleSheets>";
    private static string Pages(string shapes) => "<Pages><Page ID='0' Name='Source'><Shapes>" + shapes + "</Shapes></Page></Pages>";
    private static string Shape(string id, string attributes) => "<Shape ID='" + id + "' NameU='Node" + id + "' " + attributes
        + "><XForm><PinX>1</PinX><PinY>1</PinY><Width>2</Width><Height>1</Height></XForm><Text>Text</Text></Shape>";
    private static string Connector(string id, string attributes) => "<Shape ID='" + id + "' " + attributes
        + "><XForm1D><BeginX>1</BeginX><BeginY>1</BeginY><EndX>2</EndX><EndY>2</EndY></XForm1D><Text>Label" + id + "</Text></Shape>";
    private static VisioDocument Load(string content) => VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(
        "<VisioDocument xmlns='" + Legacy + "'>" + content + "</VisioDocument>"))).Value;
    private static VisioDocument ReopenPackage(VisioDocument document) => VisioDocument.Load(new MemoryStream(document.ToBytes()));
    private static VisioDocument ReopenLegacy(VisioDocument document) => VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value;
    private static XDocument Export(VisioDocument document) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
    private static XDocument ReadPart(VisioDocument document, string path) {
        using var archive = new ZipArchive(new MemoryStream(document.ToBytes()), ZipArchiveMode.Read);
        using var stream = archive.GetEntry(path)!.Open();
        return XDocument.Load(stream);
    }
    private static void AssertReferences(XElement element, string? line, string? fill, string? text) {
        Assert.Equal(line, (string?)element.Attribute("LineStyle"));
        Assert.Equal(fill, (string?)element.Attribute("FillStyle"));
        Assert.Equal(text, (string?)element.Attribute("TextStyle"));
    }
}
