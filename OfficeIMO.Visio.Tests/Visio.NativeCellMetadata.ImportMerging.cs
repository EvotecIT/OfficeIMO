using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioNativeCellMetadataImportMergingTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";

    [Theory]
    [InlineData("0", false)]
    [InlineData("0", true)]
    [InlineData("10", false)]
    [InlineData("10", true)]
    public void MatchedStyleCollectionsRetainAbsentDeletionAttributes(string styleId, bool sectionDeletion) {
        var destination = Load(Styles(styleId, "<Para IX='0'><HorzAlign>0</HorzAlign></Para>"));
        var source = Load(Styles(styleId, "<Para IX='0' Del='1'><HorzAlign>1</HorzAlign><BulletStr V='null'/></Para>"
            + "<Para IX='7' Del='1'><BulletStr V='null'/></Para><Scratch IX='0' Del='1'><X>0</X></Scratch>")
            + Master("<Shape ID='1'><XForm><Width>1</Width><Height>1</Height></XForm></Shape>"));
        byte[] package = source.ToBytes();
        using (var stream = new MemoryStream()) {
            stream.Write(package, 0, package.Length);
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, true)) {
                var entry = archive.GetEntry("visio/document.xml")!;
                XDocument xml;
                using (var input = entry.Open()) xml = XDocument.Load(input);
                XElement style = xml.Descendants(Modern + "StyleSheet").Single(s => (string?)s.Attribute("ID") == styleId);
                if (sectionDeletion) style.Elements(Modern + "Section").Single(s => (string?)s.Attribute("N") == "Paragraph").SetAttributeValue("Del", "1");
                style.Elements(Modern + "Section").Single(s => (string?)s.Attribute("N") == "Scratch").SetAttributeValue("Del", "1");
                entry.Delete();
                using var output = archive.CreateEntry("visio/document.xml").Open(); xml.Save(output);
            }
            package = stream.ToArray();
        }
        WithPackage(package, path => destination.ImportStencilMastersAndGet(path, new[] { "Multiple roots" }));
        foreach (var candidate in new[] { destination, Reopen(destination) }) {
            XElement style = ReadDocumentXml(candidate).Descendants(Modern + "StyleSheet").Single(s => (string?)s.Attribute("ID") == styleId);
            XElement paragraph = style.Elements(Modern + "Section").Single(s => (string?)s.Attribute("N") == "Paragraph");
            Assert.Null(paragraph.Attribute("Del"));
            XElement row = paragraph.Elements(Modern + "Row").Single(r => (string?)r.Attribute("IX") == "0");
            Assert.Null(row.Attribute("Del"));
            Assert.Equal("0", (string?)row.Elements(Modern + "Cell").Single(c => (string?)c.Attribute("N") == "HorzAlign").Attribute("V"));
            Assert.Equal("1", (string?)paragraph.Elements(Modern + "Row").Single(r => (string?)r.Attribute("IX") == "7").Attribute("Del"));
            Assert.Equal("1", (string?)style.Elements(Modern + "Section").Single(s => (string?)s.Attribute("N") == "Scratch").Attribute("Del"));
        }
    }

    [Theory]
    [InlineData("0", false)]
    [InlineData("0", true)]
    [InlineData("10", false)]
    [InlineData("10", true)]
    public void ExistingStyleHeadersRetainAbsenceWhileNewStylesRetainSourceAttributes(string styleId, bool existing) {
        string attribute = styleId == "0" ? "IsCustomNameU" : "LineStyle";
        string value = styleId == "0" ? "1" : "5";
        string sourceStyle = "<StyleSheets><StyleSheet ID='" + styleId + "' NameU='Shared' " + attribute + "='" + value
            + "'><Misc><Comment V='null'/></Misc></StyleSheet><StyleSheet ID='5' NameU='Inherited'><Line><LineColor>#123456</LineColor></Line></StyleSheet></StyleSheets>";
        var source = Load(sourceStyle + Master("<Shape ID='1'><XForm><Width>1</Width><Height>1</Height></XForm></Shape>"));
        var destination = existing ? Load(Styles(styleId, "<Misc><Comment>destination</Comment></Misc>")) : VisioDocument.Create();
        WithPackage(source, path => destination.ImportStencilMastersAndGet(path, new[] { "Multiple roots" }));
        foreach (var candidate in new[] { destination, Reopen(destination) }) {
            XElement style = ReadDocumentXml(candidate).Descendants(Modern + "StyleSheet").Single(s => (string?)s.Attribute("ID") == styleId);
            Assert.Equal(existing ? null : value, (string?)style.Attribute(attribute));
            XElement comment = Export(candidate).Descendants(Legacy + "StyleSheet").Single(s => (string?)s.Attribute("ID") == styleId)
                .Element(Legacy + "Misc")!.Element(Legacy + "Comment")!;
            Assert.Equal(existing ? "destination" : "", comment.Value);
            Assert.Equal(existing ? null : "null", (string?)comment.Attribute("V"));
        }
    }

    [Theory]
    [InlineData("0", "cell")]
    [InlineData("0", "row")]
    [InlineData("0", "root")]
    [InlineData("10", "cell")]
    [InlineData("10", "row")]
    [InlineData("10", "root")]
    public void ImportedStylesMergeNativeIdentitiesAndTransferOnlyAddedCellState(string styleId, string addition) {
        const string destinationCells = "<Misc><Comment V='null' F='Inh' Err='destination error'/></Misc>"
            + "<Para IX='0'><HorzAlign>0</HorzAlign></Para>";
        string importedCells = "<Misc><Comment F='Inh' Err='source error'>source literal</Comment>"
            + (addition == "root" ? "<ShapeKeywords V='null' F='Inh' Err='added error'/>" : "") + "</Misc>"
            + "<Para IX='0'><HorzAlign>1</HorzAlign>"
            + (addition == "cell" ? "<BulletStr V='null' F='Inh' Err='added error'/>" : "") + "</Para>"
            + (addition == "row" ? "<Para IX='7'><BulletStr V='null' F='Inh' Err='added error'/></Para>" : "")
            + "<ext:Keep xmlns:ext='urn:producer' payload='retained'/>";
        var source = Load(Styles(styleId, importedCells) + Master("<Shape ID='1'><XForm><Width>1</Width><Height>1</Height></XForm></Shape>"));
        var destination = Load(Styles(styleId, destinationCells));
        WithPackage(source, path => {
            destination.ImportStencilMastersAndGet(path, new[] { "Multiple roots" });
            // Repeated import must neither duplicate identities nor attach a source marker to a retained cell.
            destination.ImportStencilMastersAndGet(path, new[] { "Multiple roots" });
        });
        foreach (var candidate in new[] { destination, Reopen(destination) }) {
            XElement style = Export(candidate).Descendants(Legacy + "StyleSheet").Single(s => (string?)s.Attribute("ID") == styleId);
            XElement comment = Assert.Single(style.Elements(Legacy + "Misc")).Element(Legacy + "Comment")!;
            AssertState(comment, "destination error");
            XElement paragraph = style.Elements(Legacy + "Para").Single(p => (string?)p.Attribute("IX") == "0");
            Assert.Equal("0", Assert.Single(paragraph.Elements(Legacy + "HorzAlign")).Value);
            XElement added = addition == "root" ? style.Element(Legacy + "Misc")!.Element(Legacy + "ShapeKeywords")!
                : style.Elements(Legacy + "Para").Single(p => (string?)p.Attribute("IX") == (addition == "row" ? "7" : "0"))
                    .Element(Legacy + "BulletStr")!;
            AssertState(added, "added error");
            Assert.Equal(addition == "row" ? 2 : 1, style.Elements(Legacy + "Para").Count());
            Assert.Single(style.Elements(XName.Get("Keep", "urn:producer")));
        }
    }

    [Theory]
    [InlineData("same")]
    [InlineData("register")]
    [InlineData("stencil")]
    public void AdditionalMasterRootsCarryOwnAndNestedGuardedStateAcrossTransfer(string route) {
        const string first = "<Shape ID='1' NameU='Modeled'><XForm><Width>1</Width><Height>1</Height></XForm></Shape>";
        var source = Load("<FaceNames><FaceName ID='5' Name='Source family'/></FaceNames>" + Master(first + RawRoots));
        var destination = route == "same" ? source : Load("<FaceNames><FaceName ID='5' Name='Destination family'/></FaceNames>");
        VisioMaster master;
        if (route == "stencil") {
            master = null!;
            WithPackage(source, path => master = destination.ImportStencilMastersAndGet(path, new[] { "Multiple roots" }).Single());
        } else master = route == "same" ? source.GetMaster("Multiple roots") : destination.RegisterMaster(source.GetMaster("Multiple roots"));
        // The canonical ID owner must continue reserving all preserved raw root identities.
        master.Shape.Children.Add(new VisioShape("2", 0, 0, 1, 1, "inserted"));
        foreach (var candidate in new[] { destination, Reopen(destination) }) {
            XDocument xml = Export(candidate);
            XElement definition = xml.Descendants(Legacy + "Master").Single(m => (string?)m.Attribute("NameU") == "Multiple roots");
            foreach (string name in new[] { "RawRoot", "RawChild", "RawConnector" }) {
                XElement raw = definition.Descendants(Legacy + "Shape").Single(s => (string?)s.Attribute("NameU") == name);
                AssertState(raw.Element(Legacy + "Misc")!.Element(Legacy + "Comment")!, "raw error");
            }
            string[] ids = definition.Descendants(Legacy + "Shape").Select(s => ((uint)s.Attribute("ID")!).ToString()).ToArray();
            Assert.Equal(ids.Length, ids.Distinct().Count());
            XElement font = definition.Descendants(Legacy + "Font").Single();
            string expectedFont = route != "same" ? (string)xml.Descendants(Legacy + "FaceName")
                .Single(f => (string?)f.Attribute("Name") == "Source family").Attribute("ID")! : "5";
            if (route != "same") Assert.NotEqual("5", expectedFont);
            Assert.Equal(expectedFont, font.Value);
            Assert.Equal("GUARD(" + expectedFont + ")", (string?)font.Attribute("F"));
            Assert.Equal("font error", (string?)font.Attribute("Err"));
        }
    }

    [Fact]
    public void PageCopiesModelEveryNativeRootAndKeepEachRootsMetadata() {
        var document = Load("<Pages><Page ID='0' Name='Source'><Shapes>" + RawRoots + "</Shapes></Page></Pages>");
        document.DuplicatePage(document.Pages[0], "Copy");
        foreach (var candidate in new[] { document, Reopen(document) }) {
            foreach (XElement page in Export(candidate).Descendants(Legacy + "Page")) {
                foreach (string name in new[] { "RawRoot", "RawChild", "RawConnector" }) {
                    XElement shape = page.Descendants(Legacy + "Shape").Single(s => (string?)s.Attribute("NameU") == name || (string?)s.Element(Legacy + "Text") == name);
                    AssertState(shape.Element(Legacy + "Misc")!.Element(Legacy + "Comment")!, "raw error");
                }
            }
        }
    }

    private const string RawRoots = "<Shape ID='002' NameU='RawRoot' Type='Group'><XForm><Width>1</Width><Height>1</Height></XForm>"
        + "<Misc><Comment V='null' F='Inh' Err='raw error'/></Misc>"
        + "<Char IX='0'><Font F='GUARD(5)' Err='font error'>5</Font></Char>"
        + "<Shapes><Shape ID='003' NameU='RawChild'><XForm><Width>1</Width><Height>1</Height></XForm>"
        + "<Misc><Comment V='null' F='Inh' Err='raw error'/></Misc></Shape></Shapes></Shape>"
        + "<Shape ID='004' NameU='RawConnector'><XForm1D><BeginX>0</BeginX><BeginY>0</BeginY><EndX>1</EndX><EndY>1</EndY></XForm1D>"
        + "<Misc><Comment V='null' F='Inh' Err='raw error'/></Misc><Text>RawConnector</Text></Shape>";
    private static string Styles(string id, string cells) => "<StyleSheets><StyleSheet ID='" + id + "' NameU='Shared'>" + cells + "</StyleSheet></StyleSheets>";
    private static string Master(string shapes) => "<Masters><Master ID='0' NameU='Multiple roots'><Shapes>" + shapes + "</Shapes></Master></Masters>";
    private static VisioDocument Load(string content) => VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(
        "<VisioDocument xmlns='" + Legacy + "'>" + content + "</VisioDocument>")), VisioPackageType.Stencil).Value;
    private static VisioDocument Reopen(VisioDocument document) => VisioDocument.Load(new MemoryStream(document.ToBytes()));
    private static XDocument Export(VisioDocument document) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
    private static XDocument ReadDocumentXml(VisioDocument document) {
        using var archive = new ZipArchive(new MemoryStream(document.ToBytes()), ZipArchiveMode.Read);
        using var stream = archive.GetEntry("visio/document.xml")!.Open(); return XDocument.Load(stream);
    }
    private static void AssertState(XElement cell, string error) {
        Assert.Equal("", cell.Value); Assert.Equal("null", (string?)cell.Attribute("V"));
        Assert.Equal("Inh", (string?)cell.Attribute("F")); Assert.Equal(error, (string?)cell.Attribute("Err"));
    }
    private static void WithPackage(VisioDocument document, Action<string> action) => WithPackage(document.ToBytes(), action);
    private static void WithPackage(byte[] package, Action<string> action) {
        string path = Path.Combine(AppContext.BaseDirectory, "native-import-" + Guid.NewGuid().ToString("N") + ".vssx");
        try { File.WriteAllBytes(path, package); action(path); }
        finally { File.Delete(path); }
    }
}
