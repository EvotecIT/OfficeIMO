using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioNativeCellMetadataStyleIdentityTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";

    [Theory]
    [InlineData("10", "010")]
    [InlineData("010", "10")]
    [InlineData("0", "00")]
    [InlineData("00", "0")]
    public void NumericStyleAliasesMergeOnceAndKeepDestinationAndAddedCellMetadata(string destinationId, string sourceId) {
        var destination = Load(Style(destinationId, "<Misc><Comment V='null' F='Inh' Err='destination error'/></Misc>"));
        var source = Load(Style(sourceId, "<Misc><Comment V='null' F='Inh' Err='source error'/>"
            + "<ShapeKeywords V='null' F='Inh' Err='added error'/></Misc>") + Master);
        WithPackage(source, path => {
            destination.ImportStencilMastersAndGet(path, new[] { "Tiny" });
            destination.ImportStencilMastersAndGet(path, new[] { "Tiny" });
        });
        foreach (var candidate in new[] { destination, Reopen(destination) }) {
            XElement style = NumericStyle(Export(candidate), Legacy, destinationId);
            XElement misc = Assert.Single(style.Elements(Legacy + "Misc"));
            AssertState(Assert.Single(misc.Elements(Legacy + "Comment")), "destination error");
            AssertState(Assert.Single(misc.Elements(Legacy + "ShapeKeywords")), "added error");
            NumericStyle(ReadDocumentXml(candidate), Modern, destinationId);
        }
        AssertState(NumericStyle(Export(source), Legacy, sourceId).Element(Legacy + "Misc")!.Element(Legacy + "Comment")!, "source error");
    }

    [Theory]
    [InlineData("00")]
    [InlineData("010")]
    public void NewNumericStylesRetainSourceHeaderRowsAndGuardedCells(string sourceId) {
        var source = Load(Style(sourceId, "<User ID='7' NameU='Fresh' Del='1'><Value V='null' F='Inh' Err='new error'/></User>", " IsCustomNameU='1'") + Master);
        // Saving a leading-zero generated style must not synthesize a second numeric identity.
        NumericStyle(ReadDocumentXml(source), Modern, sourceId);
        var destination = VisioDocument.Create();
        WithPackage(source, path => destination.ImportStencilMastersAndGet(path, new[] { "Tiny" }));
        foreach (var candidate in new[] { destination, Reopen(destination) }) {
            XElement style = NumericStyle(Export(candidate), Legacy, sourceId);
            Assert.Equal("1", (string?)style.Attribute("IsCustomNameU"));
            XElement row = Assert.Single(style.Elements(Legacy + "User"));
            Assert.Equal("7", (string?)row.Attribute("ID"));
            Assert.Equal("Fresh", (string?)row.Attribute("NameU"));
            Assert.Equal("1", (string?)row.Attribute("Del"));
            AssertState(row.Element(Legacy + "Value")!, "new error");
        }
    }

    [Theory]
    [InlineData("0")]
    [InlineData("10")]
    public void NamedStyleRowsMergeByUniversalNameAndRebaseOnlyAddedCellMetadata(string styleId) {
        var destination = Load(Style(styleId, "<User ID='0' NameU='Flag'><Value F='Inh'>destination</Value></User>"));
        var source = Load(Style(styleId, "<User ID='7' NameU='Flag' Del='1'><Value V='null' F='Inh' Err='source error'/>"
            + "<Prompt V='null' F='Inh' Err='added error'/></User>"
            + "<User ID='8' NameU='Fresh' Del='1'><Value V='null' F='Inh' Err='new error'/></User>") + Master);
        WithPackage(source, path => {
            destination.ImportStencilMastersAndGet(path, new[] { "Tiny" });
            destination.ImportStencilMastersAndGet(path, new[] { "Tiny" });
        });
        foreach (var candidate in new[] { destination, Reopen(destination) }) {
            XElement style = NumericStyle(Export(candidate), Legacy, styleId);
            XElement flag = Assert.Single(style.Elements(Legacy + "User"), r => (string?)r.Attribute("NameU") == "Flag");
            Assert.Equal("0", (string?)flag.Attribute("ID"));
            Assert.Null(flag.Attribute("Del"));
            XElement value = Assert.Single(flag.Elements(Legacy + "Value"));
            Assert.Equal("destination", value.Value); Assert.Null(value.Attribute("V")); Assert.Null(value.Attribute("Err"));
            AssertState(Assert.Single(flag.Elements(Legacy + "Prompt")), "added error");
            XElement fresh = Assert.Single(style.Elements(Legacy + "User"), r => (string?)r.Attribute("NameU") == "Fresh");
            Assert.Equal("8", (string?)fresh.Attribute("ID")); Assert.Equal("1", (string?)fresh.Attribute("Del"));
            AssertState(fresh.Element(Legacy + "Value")!, "new error");
        }
        XElement original = NumericStyle(Export(source), Legacy, styleId).Elements(Legacy + "User").Single(r => (string?)r.Attribute("NameU") == "Flag");
        Assert.Equal("7", (string?)original.Attribute("ID")); AssertState(original.Element(Legacy + "Value")!, "source error");
    }

    private const string Master = "<Masters><Master ID='0' NameU='Tiny'><Shapes><Shape ID='1'><XForm><Width>1</Width><Height>1</Height></XForm></Shape></Shapes></Master></Masters>";
    private static string Style(string id, string cells, string attributes = "") => "<StyleSheets><StyleSheet ID='" + id + "' NameU='Shared'" + attributes + ">" + cells + "</StyleSheet></StyleSheets>";
    private static VisioDocument Load(string content) => VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(
        "<VisioDocument xmlns='" + Legacy + "'>" + content + "</VisioDocument>")), VisioPackageType.Stencil).Value;
    private static VisioDocument Reopen(VisioDocument document) => VisioDocument.Load(new MemoryStream(document.ToBytes()));
    private static XDocument Export(VisioDocument document) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
    private static XElement NumericStyle(XDocument xml, XNamespace ns, string id) => Assert.Single(xml.Descendants(ns + "StyleSheet"),
        s => ulong.Parse((string)s.Attribute("ID")!, CultureInfo.InvariantCulture) == ulong.Parse(id, CultureInfo.InvariantCulture));
    private static XDocument ReadDocumentXml(VisioDocument document) {
        using var archive = new ZipArchive(new MemoryStream(document.ToBytes()), ZipArchiveMode.Read);
        using var stream = archive.GetEntry("visio/document.xml")!.Open(); return XDocument.Load(stream);
    }
    private static void AssertState(XElement cell, string error) {
        Assert.Equal("", cell.Value); Assert.Equal("null", (string?)cell.Attribute("V"));
        Assert.Equal("Inh", (string?)cell.Attribute("F")); Assert.Equal(error, (string?)cell.Attribute("Err"));
    }
    private static void WithPackage(VisioDocument document, Action<string> action) {
        string path = Path.Combine(AppContext.BaseDirectory, "native-identity-" + Guid.NewGuid().ToString("N") + ".vssx");
        try { File.WriteAllBytes(path, document.ToBytes()); action(path); }
        finally { File.Delete(path); }
    }
}
