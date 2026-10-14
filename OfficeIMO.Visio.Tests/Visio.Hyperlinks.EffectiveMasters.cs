using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioHyperlinkEffectiveMasterTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImplicitShapeAndDynamicConnectorNamesStayStableAcrossSelectionAndPageCopies(bool imported) {
        VisioDocument document = VisioDocument.Create();
        document.UseMastersByDefault = true;
        if (imported) {
            VisioDocument stencil = LoadFixture(includePage: false);
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vssx");
            try {
                File.WriteAllBytes(path, stencil.ToBytes());
                document.ImportStencilMasters(path);
            } finally { File.Delete(path); }
        } else {
            var rectangle = new VisioShape("1", 1, 1, 2, 1, "") { NameU = "Rectangle" };
            rectangle.AddHyperlink("https://example.org/master-shape");
            rectangle.Children.Add(new VisioShape("2", 1, 1, 1, 1, "Child"));
            document.RegisterMaster("Rectangle", rectangle);
            var dynamic = new VisioShape("1", 0, 0, 2, 0, "");
            dynamic.AddHyperlink("https://example.org/master-connector");
            document.RegisterMaster("Dynamic connector", dynamic);
        }
        VisioPage page = document.AddPage("Source");
        VisioShape first = page.AddRectangle(2, 2, 2, 1, "First");
        VisioShape second = page.AddRectangle(6, 2, 2, 1, "Second");
        VisioConnector connector = page.AddConnector(first, second);
        first.AddHyperlink("https://example.org/local-shape");
        connector.AddHyperlink("https://example.org/local-connector");
        Assert.Null(first.Master);
        page.SelectShapes(shape => ReferenceEquals(shape, first) || ReferenceEquals(shape, second)).Duplicate();
        document.DuplicatePage(page, "Copy");
        foreach (VisioDocument saved in Reopened(document)) {
            AssertLocalNames(saved, new Dictionary<string, string> {
                ["https://example.org/local-shape"] = "Row_2", ["https://example.org/local-connector"] = "Row_2"
            }, 8);
            using var archive = new ZipArchive(new MemoryStream(saved.ToBytes()), ZipArchiveMode.Read);
            XElement[] inherited = ReadParts(archive, "visio/masters/").SelectMany(part => part.Descendants(Modern + "Section"))
                .Where(section => (string?)section.Attribute("N") == "Hyperlink").Elements(Modern + "Row").ToArray();
            Assert.Equal(2, inherited.Length);
            Assert.All(inherited, row => Assert.Equal("Row_1", (string?)row.Attribute("N")));
            if (imported) Assert.All(inherited, row => Assert.Equal("7", (string?)row.Attribute("IX")));
        }
        Assert.Null(first.Master);
        Assert.Null(first.Hyperlinks.Single().RowName);
        Assert.Null(connector.Hyperlinks.Single().RowName);
    }

    [Fact]
    public void PreservedDynamicConnectorMasterStillReservesNamesWhenAutomaticMastersAreDisabled() {
        VisioDocument document = LoadFixture(includePage: true);
        document.UseMastersByDefault = false;
        VisioPage page = document.Pages[0];
        VisioConnector connector = page.Connectors.Single();
        connector.Hyperlinks.Clear();
        connector.AddHyperlink("https://example.org/local-connector");
        page.SelectShapes(_ => true).Duplicate();
        document.DuplicatePage(page, "Copy");
        foreach (VisioDocument saved in Reopened(document))
            AssertLocalNames(saved, new Dictionary<string, string> { ["https://example.org/local-connector"] = "Row_2" }, 4);
        Assert.Null(connector.Hyperlinks.Single().RowName);
    }

    [Fact]
    public void ImplicitMasterCopiesPreserveUntouchedLocalMetadataAndExplicitOverrides() {
        VisioDocument document = LoadFixture(includePage: true);
        document.UseMastersByDefault = true;
        VisioPage page = document.Pages[0];
        VisioShape shape = page.Shapes[0];
        VisioConnector connector = page.Connectors.Single();
        foreach (IList<VisioHyperlink> rows in new[] { shape.Hyperlinks, connector.Hyperlinks })
            rows.Insert(0, new VisioHyperlink("https://example.org/local-added"));
        shape.AddHyperlink("https://example.org/local-override").RowName = "Row_1";
        connector.AddHyperlink("https://example.org/local-override").RowName = "Row_1";
        page.SelectShapes(_ => true).Duplicate();
        document.DuplicatePage(page, "Copy");
        foreach (VisioDocument saved in Reopened(document)) {
            AssertLocalNames(saved, new Dictionary<string, string> {
                ["https://example.org/local-added"] = "Row_3", ["https://example.org/local-existing"] = "Row_2",
                ["https://example.org/local-override"] = "Row_1"
            }, 24);
            XDocument legacy = XDocument.Load(new MemoryStream(saved.ToLegacyXmlResult().Value));
            XElement[] unchanged = legacy.Descendants(Legacy + "Page").Descendants(Legacy + "Hyperlink")
                .Where(row => row.Element(Legacy + "Address")!.Value == "https://example.org/local-existing").ToArray();
            Assert.Equal(8, unchanged.Length);
            foreach (XElement row in unchanged) {
                XElement description = row.Element(Legacy + "Description")!;
                Assert.Equal("null", (string?)description.Attribute("V"));
                Assert.Equal("description error", (string?)description.Attribute("Err"));
                Assert.Equal("STR", (string?)description.Attribute("Unit"));
                Assert.Equal("\"https://example.org/local-existing\"", (string?)row.Element(Legacy + "Address")!.Attribute("F"));
            }
        }
        Assert.Null(shape.Master);
        Assert.Null(shape.Hyperlinks[0].RowName);
        Assert.Null(connector.Hyperlinks[0].RowName);
    }

    [Fact]
    public void NewMasterInstancesAndNestedCopiesKeepTheirCorrespondingMasterRowIdentities() {
        VisioDocument document = VisioDocument.Create();
        document.UseMastersByDefault = true;
        var blueprint = new VisioShape("1", 1, 1, 2, 1, "") { NameU = "Links" };
        blueprint.AddHyperlink("https://example.org/master-root");
        var childBlueprint = new VisioShape("2", 1, 1, 1, 1, "") { NameU = "Rectangle" };
        childBlueprint.AddHyperlink("https://example.org/master-child").RowName = "Row_2";
        blueprint.Children.Add(childBlueprint);
        VisioMaster master = document.RegisterMaster("Links", blueprint);
        var unrelated = new VisioShape("1", 1, 1, 2, 1, "");
        unrelated.AddHyperlink("https://example.org/other-master");
        document.RegisterMaster("Rectangle", unrelated);
        VisioPage page = document.AddPage("Instances");
        VisioShape root = page.AddShape("10", master, 3, 3, 2, 1);
        VisioShape child = root.Children.Single();
        Assert.Equal("Row_1", Assert.Single(root.Hyperlinks).RowName);
        Assert.Equal("Row_2", Assert.Single(child.Hyperlinks).RowName);
        root.Hyperlinks.Clear();
        child.Hyperlinks.Clear();
        root.AddHyperlink("https://example.org/local-root");
        child.AddHyperlink("https://example.org/local-child");
        page.SelectShapes(shape => ReferenceEquals(shape, root)).Duplicate();
        document.DuplicatePage(page, "Copy");
        foreach (VisioDocument saved in Reopened(document))
            AssertLocalNames(saved, new Dictionary<string, string> {
                ["https://example.org/local-root"] = "Row_2", ["https://example.org/local-child"] = "Row_1"
            }, 8);
        Assert.Null(blueprint.Hyperlinks.Single().RowName);
        Assert.Null(root.Hyperlinks.Single().RowName);
        Assert.Null(child.Hyperlinks.Single().RowName);
    }

    private static VisioDocument LoadFixture(bool includePage) {
        string inherited = "<Hyperlink ID='7' NameU='Row_1'><Address Unit='STR'>https://example.org/master-";
        string local = "<Hyperlink NameU='Row_2'><Description V='null' Unit='STR' Err='description error'/>"
            + "<Address Unit='STR' F='&quot;https://example.org/local-existing&quot;'>https://example.org/local-existing</Address></Hyperlink>";
        string xml = "<VisioDocument xmlns='" + Legacy + "'><Masters><Master ID='1' NameU='Rectangle'><Shapes><Shape ID='1' NameU='Rectangle'>"
            + "<XForm><Width>2</Width><Height>1</Height></XForm>" + inherited + "shape</Address></Hyperlink></Shape></Shapes></Master>"
            + "<Master ID='2' NameU='Dynamic connector'><Shapes><Shape ID='1' NameU='Dynamic connector'>"
            + "<XForm><Width>2</Width><Height>0</Height></XForm>" + inherited + "connector</Address></Hyperlink></Shape></Shapes></Master></Masters>";
        if (includePage) xml += "<Pages><Page ID='0' Name='Source'><Shapes><Shape ID='1' NameU='Rectangle'><XForm><Width>2</Width><Height>1</Height></XForm>"
            + local + "</Shape><Shape ID='2'><XForm><Width>2</Width><Height>1</Height></XForm></Shape>"
            + "<Shape ID='3' NameU='Dynamic connector' Master='2'><XForm1D><BeginX>0</BeginX><BeginY>0</BeginY><EndX>2</EndX><EndY>0</EndY></XForm1D>"
            + local + "</Shape></Shapes><Connects><Connect FromSheet='3' FromCell='BeginX' ToSheet='1' ToCell='PinX'/>"
            + "<Connect FromSheet='3' FromCell='EndX' ToSheet='2' ToCell='PinX'/></Connects></Page></Pages>";
        xml += "</VisioDocument>";
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml)), includePage ? VisioPackageType.Drawing : VisioPackageType.Stencil).Value;
    }

    private static IEnumerable<VisioDocument> Reopened(VisioDocument document) {
        yield return document;
        yield return VisioDocument.Load(new MemoryStream(document.ToBytes()));
        yield return VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value;
    }

    private static void AssertLocalNames(VisioDocument document, IReadOnlyDictionary<string, string> names, int count) {
        XDocument legacy = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
        XElement[] rows = legacy.Descendants(Legacy + "Page").Descendants(Legacy + "Hyperlink")
            .Where(row => names.ContainsKey(row.Element(Legacy + "Address")!.Value)).ToArray();
        Assert.Equal(count, rows.Length);
        Assert.All(rows, row => Assert.Equal(names[row.Element(Legacy + "Address")!.Value], (string?)row.Attribute("NameU")));
        using var archive = new ZipArchive(new MemoryStream(document.ToBytes()), ZipArchiveMode.Read);
        XElement[] native = ReadParts(archive, "visio/pages/").SelectMany(part => part.Descendants(Modern + "Section"))
            .Where(section => (string?)section.Attribute("N") == "Hyperlink").Elements(Modern + "Row")
            .Where(row => names.ContainsKey(row.Elements(Modern + "Cell").Single(cell => (string?)cell.Attribute("N") == "Address").Attribute("V")!.Value)).ToArray();
        Assert.Equal(count, native.Length);
        Assert.All(native, row => Assert.Equal(names[row.Elements(Modern + "Cell").Single(cell => (string?)cell.Attribute("N") == "Address").Attribute("V")!.Value], (string?)row.Attribute("N")));
    }

    private static IEnumerable<XDocument> ReadParts(ZipArchive archive, string prefix) {
        foreach (ZipArchiveEntry entry in archive.Entries.Where(entry => entry.FullName.StartsWith(prefix, StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal))) {
            using Stream stream = entry.Open();
            yield return XDocument.Load(stream);
        }
    }
}
