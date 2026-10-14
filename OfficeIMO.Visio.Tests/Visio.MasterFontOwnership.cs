using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioMasterFontOwnershipTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";
    private const string Transform = "<XForm><PinX>1</PinX><PinY>1</PinY><Width>2</Width><Height>1</Height></XForm>";
    private const string Master = "<Masters><Master ID='8' NameU='Shared fonts'><Shapes><Shape ID='1'>" + Transform
        + "<Char IX='0'><Font F='GUARD(5)' Err='source font error'>5</Font><Size>0.125</Size></Char>"
        + "<Text><cp IX='0'/>Master label</Text></Shape></Shapes></Master></Masters>";
    private const string Faces = "<FaceNames><FaceName ID='5' Name='Source family' CharSets='0'/>"
        + "<FaceName ID='9' Name='Source family' CharSets='128'/></FaceNames>";

    [Fact]
    public void RegisteredMasterKeepsIdentityAndSourceFontsAcrossMultipleDestinationSaves() {
        VisioDocument source = Load(Faces + Master);
        VisioMaster shared = source.GetMaster("Shared fonts");
        XDocument sourceBefore = LegacyXml(source);
        VisioDocument first = Destination("First destination"), second = Destination("Second destination");
        Assert.Same(shared, first.RegisterMaster(shared));
        Assert.Same(shared, second.RegisterMaster(shared));
        first.AddPage("First").AddShape("1", shared, 1, 1, 2, 1);
        second.AddPage("Second").AddShape("1", shared, 1, 1, 2, 1);

        for (int repeat = 0; repeat < 2; repeat++) {
            foreach (VisioDocument destination in new[] { first, second }) {
                Assert.Same(shared, destination.GetMaster("Shared fonts"));
                AssertSourceFamily(destination);
                byte[] bytes = destination.ToBytes();
                XDocument definition = ReadPart(bytes, "visio/masters/master1.xml");
                XElement font = definition.Descendants(Modern + "Cell").Single(c => (string?)c.Attribute("N") == "Font");
                Assert.Equal("Source family", (string?)font.Attribute("V"));
                Assert.Equal("GUARD(FONT(\"Source family\"))", (string?)font.Attribute("F"));
                VisioDocument reopened = VisioDocument.Load(new MemoryStream(bytes));
                AssertSourceFamily(reopened);
                Assert.Equal("Source family", reopened.GetMaster("Shared fonts").Shape.TextStyle!.FontFamily);
                Assert.True(XNode.DeepEquals(sourceBefore, LegacyXml(source)));
            }
        }
    }

    [Theory]
    [InlineData("Source family")]
    [InlineData("Edited family")]
    public void ExplicitMasterFamilyAssignmentsReachOutputAndClearProducerError(string family) {
        VisioDocument source = Load(Faces + Master);
        VisioMaster master = source.GetMaster("Shared fonts");
        master.Shape.TextStyle!.FontFamily = family;
        VisioDocument destination = Destination("Destination family");
        Assert.Same(master, destination.RegisterMaster(master));
        destination.AddPage("Instance").AddShape("1", master, 1, 1, 2, 1);
        foreach (VisioDocument candidate in new[] { destination, VisioDocument.Load(new MemoryStream(destination.ToBytes())) }) {
            XDocument xml = LegacyXml(candidate);
            XElement font = xml.Descendants(Legacy + "Master").Descendants(Legacy + "Font").Single();
            string id = font.Value;
            Assert.Equal(family, (string?)xml.Descendants(Legacy + "FaceName").Single(face => (string?)face.Attribute("ID") == id).Attribute("Name"));
            Assert.Null(font.Attribute("F"));
            Assert.Null(font.Attribute("Err"));
            AssertFamily(candidate, family);
        }
    }

    [Fact]
    public void RegisteredRawMasterRootsRebindAllFontKindsAndKeepSentinelsCharsetAndErrors() {
        const string raw = "<Shape ID='2' NameU='Raw fonts' Type='Group'>" + Transform
            + "<Char IX='3'><Font F='GUARD(9)' Err='font error'>9</Font>"
            + "<AsianFont F='GUARD(9)' Err='asian error'>9</AsianFont><ComplexScriptFont F='9' Err='complex error'>9</ComplexScriptFont></Char>"
            + "<Char IX='4'><Font>5</Font><AsianFont>0</AsianFont><ComplexScriptFont></ComplexScriptFont></Char>"
            + "<Para IX='3'><BulletFont F='GUARD(9)' Err='bullet error'>9</BulletFont></Para><Para IX='4'><BulletFont>0</BulletFont></Para>"
            + "<Shapes><Shape ID='3' NameU='Nested fonts'>" + Transform
            + "<Char IX='0'><Font F='GUARD(5)' Err='nested error'>5</Font></Char><Text>Nested</Text></Shape></Shapes></Shape>";
        string masterXml = Master.Replace("</Shapes></Master>", raw + "</Shapes></Master>");
        VisioDocument source = Load(Faces + masterXml);
        XDocument before = LegacyXml(source);
        VisioDocument destination = Destination("Destination family", VisioPackageType.Stencil);
        destination.RegisterMaster(source.GetMaster("Shared fonts"));
        foreach (VisioDocument candidate in new[] { destination, VisioDocument.Load(new MemoryStream(destination.ToBytes())) }) {
            XDocument xml = LegacyXml(candidate);
            XElement definition = xml.Descendants(Legacy + "Master").Single();
            XElement root = definition.Descendants(Legacy + "Shape").Single(shape => (string?)shape.Attribute("NameU") == "Raw fonts");
            string variant = (string)xml.Descendants(Legacy + "FaceName").Single(face =>
                (string?)face.Attribute("Name") == "Source family" && (string?)face.Attribute("CharSets") == "128").Attribute("ID")!;
            Assert.NotEqual("0", variant);
            Assert.NotEqual("5", variant);
            XElement row = root.Elements(Legacy + "Char").Single(charRow => (string?)charRow.Attribute("IX") == "3");
            foreach (var value in new[] { ("Font", "font error"), ("AsianFont", "asian error"), ("ComplexScriptFont", "complex error") }) {
                XElement cell = row.Element(Legacy + value.Item1)!;
                Assert.Equal(variant, cell.Value);
                Assert.Equal(value.Item2, (string?)cell.Attribute("Err"));
                Assert.Equal(value.Item1 == "ComplexScriptFont" ? variant : "GUARD(" + variant + ")", (string?)cell.Attribute("F"));
            }
            XElement bullet = root.Elements(Legacy + "Para").Single(para => (string?)para.Attribute("IX") == "3").Element(Legacy + "BulletFont")!;
            Assert.Equal(variant, bullet.Value);
            Assert.Equal("GUARD(" + variant + ")", (string?)bullet.Attribute("F"));
            Assert.Equal("bullet error", (string?)bullet.Attribute("Err"));
            XElement sentinel = root.Elements(Legacy + "Char").Single(charRow => (string?)charRow.Attribute("IX") == "4");
            Assert.Equal("0", sentinel.Element(Legacy + "AsianFont")!.Value);
            Assert.Equal("", sentinel.Element(Legacy + "ComplexScriptFont")!.Value);
            Assert.Equal("0", root.Elements(Legacy + "Para").Single(para => (string?)para.Attribute("IX") == "4").Element(Legacy + "BulletFont")!.Value);
            XElement nested = definition.Descendants(Legacy + "Shape").Single(shape => (string?)shape.Attribute("NameU") == "Nested fonts");
            string ordinary = (string)xml.Descendants(Legacy + "FaceName").Single(face =>
                (string?)face.Attribute("Name") == "Source family" && (string?)face.Attribute("CharSets") == "0").Attribute("ID")!;
            Assert.Equal(ordinary, nested.Element(Legacy + "Char")!.Element(Legacy + "Font")!.Value);
            Assert.Equal("nested error", (string?)nested.Element(Legacy + "Char")!.Element(Legacy + "Font")!.Attribute("Err"));
        }
        Assert.True(XNode.DeepEquals(before, LegacyXml(source)));
    }

    [Fact]
    public void RegisteredLegacyFontsTableKeepsDistinctSourceCharsetEntries() {
        const string fonts = "<Fonts><FontEntry ID='5' Name='Source family' CharSet='0' PitchAndFamily='34'/>"
            + "<FontEntry ID='9' Name='Source family' CharSet='128' PitchAndFamily='49'/></Fonts>";
        VisioDocument source = Load(fonts + Master);
        VisioDocument destination = Load("<Fonts><FontEntry ID='5' Name='Destination family' CharSet='0'/></Fonts>");
        destination.RegisterMaster(source.GetMaster("Shared fonts"));
        foreach (VisioDocument candidate in new[] { destination, VisioDocument.Load(new MemoryStream(destination.ToBytes())) }) {
            XDocument xml = LegacyXml(candidate);
            foreach (string charset in new[] { "0", "128" }) {
                XElement entry = xml.Descendants(Legacy + "FontEntry").Single(font =>
                    (string?)font.Attribute("Name") == "Source family" && (string?)font.Attribute("CharSet") == charset);
                XElement face = xml.Descendants(Legacy + "FaceName").Single(value => (string?)value.Attribute("ID") == (string?)entry.Attribute("ID"));
                Assert.Equal("Source family", (string?)face.Attribute("Name"));
                Assert.Equal(charset, (string?)face.Attribute("CharSets"));
                Assert.Equal(charset == "0" ? "34" : "49", (string?)entry.Attribute("PitchAndFamily"));
            }
        }
    }

    [Fact]
    public void NewModeledMasterChildKeepsDestinationFontIdentity() {
        VisioDocument source = Load(Faces + Master);
        VisioMaster master = source.GetMaster("Shared fonts");
        master.Shape.Children.Add(new VisioShape("2", 1, 1, 1, 1, "Authored child") {
            TextStyle = new VisioTextStyle { FontFamily = "Destination family", Size = 10 }
        });
        VisioDocument destination = Destination("Destination family", VisioPackageType.Stencil);
        destination.RegisterMaster(master);
        foreach (VisioDocument candidate in new[] { destination, VisioDocument.Load(new MemoryStream(destination.ToBytes())) }) {
            XDocument xml = LegacyXml(candidate);
            XElement child = xml.Descendants(Legacy + "Master").Descendants(Legacy + "Shape")
                .Single(shape => (string?)shape.Element(Legacy + "Text") == "Authored child");
            string id = child.Element(Legacy + "Char")!.Element(Legacy + "Font")!.Value;
            Assert.Equal("Destination family", (string?)xml.Descendants(Legacy + "FaceName")
                .Single(face => (string?)face.Attribute("ID") == id).Attribute("Name"));
        }
    }

    [Fact]
    public void PaddedSourceFontIdentifiersResolveInRegisteredMasterOutputAndRendering() {
        VisioDocument source = Load(Faces.Replace("ID='5'", "ID='005'") + Master.Replace("GUARD(5)", "GUARD(005)").Replace(">5</Font>", ">005</Font>"));
        VisioDocument destination = Destination("Destination family");
        VisioMaster master = source.GetMaster("Shared fonts");
        destination.RegisterMaster(master);
        destination.AddPage("Instance").AddShape("1", master, 1, 1, 2, 1);
        foreach (VisioDocument candidate in new[] { destination, VisioDocument.Load(new MemoryStream(destination.ToBytes())) }) {
            AssertSourceFamily(candidate);
            XDocument xml = LegacyXml(candidate);
            string id = xml.Descendants(Legacy + "Master").Descendants(Legacy + "Font").Single().Value;
            Assert.Equal("Source family", (string?)xml.Descendants(Legacy + "FaceName")
                .Single(face => (string?)face.Attribute("ID") == id).Attribute("Name"));
        }
    }

    private static void AssertSourceFamily(VisioDocument document) => AssertFamily(document, "Source family");
    private static void AssertFamily(VisioDocument document, string family) {
        VisioPage page = Assert.Single(document.Pages);
        VisioShape shape = Assert.Single(page.Shapes);
        VisioRichTextProjection projection = Assert.IsType<VisioRichTextProjection>(VisioRichTextProjection.Create(page, shape, 72, CancellationToken.None));
        Assert.All(projection.Runs, run => Assert.Equal(family, run.FontFamily));
        Assert.Contains(family, page.ToSvg(), StringComparison.Ordinal);
    }
    private static VisioDocument Destination(string family, VisioPackageType type = VisioPackageType.Drawing) => Load(
        "<FaceNames><FaceName ID='5' Name='" + family + "' CharSets='0'/></FaceNames>", type);
    private static VisioDocument Load(string body, VisioPackageType type = VisioPackageType.Stencil) => VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(
        "<VisioDocument xmlns='" + Legacy + "'>" + body + "</VisioDocument>")), type).Value;
    private static XDocument LegacyXml(VisioDocument document) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
    private static XDocument ReadPart(byte[] bytes, string path) {
        using var package = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        using var input = package.GetEntry(path)!.Open();
        return XDocument.Load(input);
    }
}
