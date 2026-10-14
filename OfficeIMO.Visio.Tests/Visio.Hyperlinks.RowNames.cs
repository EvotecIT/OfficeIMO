using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioHyperlinkRowNameTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";

    [Fact]
    public void AutomaticNamesReserveEarlierLaterAndAlreadyAllocatedNamesForShapesConnectorsAndMasters() {
        var document = VisioDocument.Create();
        document.UseMastersByDefault = false;
        VisioPage page = document.AddPage("Links");
        VisioShape shape = page.AddRectangle(2, 2, 2, 1);
        VisioShape target = page.AddRectangle(6, 2, 2, 1);
        VisioConnector connector = page.AddConnector(shape, target);
        VisioMaster master = document.RegisterMaster("Links", new VisioShape("10", 1, 1, 2, 1, "Master"));
        master.Shape.Children.Add(new VisioShape("11", 1, 1, 1, 1, "Child"));
        target.Master = master;
        foreach (Func<string, VisioHyperlink> add in new Func<string, VisioHyperlink>[] {
                     address => shape.AddHyperlink(address), address => connector.AddHyperlink(address),
                     address => master.Shape.AddHyperlink(address)
                 }) {
            string?[] names = { "Row_2", null, "Row_3", null, "Manual", null, "Row_1" };
            for (int i = 0; i < names.Length; i++) add("https://example.org/" + i).RowName = names[i];
        }

        VisioPage copy = document.DuplicatePage(page, "Copy");
        string?[] expected = { "Row_2", "Row_4", "Row_3", "Row_5", "Manual", "Row_6", "Row_1" };
        AssertNames(document, expected, 5);
        AssertNames(document, expected, 5);
        foreach (IList<VisioHyperlink> rows in new[] {
                     shape.Hyperlinks, connector.Hyperlinks, master.Shape.Hyperlinks
                 }) {
            Assert.Null(rows[1].RowName);
            Assert.Null(rows[3].RowName);
            Assert.Null(rows[5].RowName);
        }
        Assert.Equal(expected, copy.Shapes[0].Hyperlinks.Select(row => row.RowName));
        Assert.Equal(expected, copy.Connectors.Single().Hyperlinks.Select(row => row.RowName));
        AssertNames(VisioDocument.Load(new MemoryStream(document.ToBytes())), expected, 5);
        AssertNames(VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value, expected, 5);
    }

    [Fact]
    public void CollisionFreeDefaultNamesAndDeliberateExplicitDuplicatesStayUnchanged() {
        var document = VisioDocument.Create();
        document.UseMastersByDefault = false;
        VisioShape shape = document.AddPage("Defaults").AddRectangle(1, 1, 2, 1);
        shape.AddHyperlink("https://example.org/one");
        shape.AddHyperlink("https://example.org/two");
        AssertNames(document, new[] { "Row_1", "Row_2" }, 1);
        Assert.All(shape.Hyperlinks, row => Assert.Null(row.RowName));

        shape.Hyperlinks[0].RowName = "Repeated";
        shape.Hyperlinks[1].RowName = "Repeated";
        shape.AddHyperlink("https://example.org/three");
        AssertNames(document, new[] { "Repeated", "Repeated", "Row_3" }, 1);
        Assert.Null(shape.Hyperlinks[2].RowName);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LoadedRowsKeepIdentityAndCellMetadataAfterRemovalReorderCopyAndNewNameAllocation(bool indexed) {
        string rows = "<Hyperlink ID='7' NameU='Row_3' Name='Third'><Address Unit='STR'>https://example.org/third</Address></Hyperlink>"
            + "<Hyperlink ID='8' NameU='Discard'><Address>https://example.org/discard</Address></Hyperlink>"
            + "<Hyperlink " + (indexed ? "ID='9' " : "") + "NameU='Row_1' Name='Localized first'><Description V='null' Unit='STR' Err='description error'/>"
            + "<Address Unit='STR' F='&quot;https://example.org/first&quot;' Err='#VALUE!'>https://example.org/first</Address></Hyperlink>"
            + "<Hyperlink ID='12'><Address Unit='STR'>https://example.org/unnamed-source</Address></Hyperlink>";
        string xml = "<VisioDocument xmlns='" + Legacy + "'><Masters><Master ID='1' NameU='Links'><Shapes><Shape ID='10'>"
            + "<XForm><Width>2</Width><Height>1</Height></XForm>" + rows + "</Shape></Shapes></Master></Masters>"
            + "<Pages><Page ID='0' Name='Links'><Shapes><Shape ID='1'><XForm><Width>2</Width><Height>1</Height></XForm>" + rows + "</Shape>"
            + "<Shape ID='2' Master='1'><XForm><Width>2</Width><Height>1</Height></XForm></Shape>"
            + "<Shape ID='3'><XForm1D><BeginX>0</BeginX><BeginY>0</BeginY><EndX>2</EndX><EndY>0</EndY></XForm1D>" + rows + "</Shape></Shapes>"
            + "<Connects><Connect FromSheet='3' FromCell='BeginX' ToSheet='1' ToCell='PinX'/>"
            + "<Connect FromSheet='3' FromCell='EndX' ToSheet='2' ToCell='PinX'/></Connects></Page></Pages></VisioDocument>";
        VisioDocument document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml))).Value;
        foreach (IList<VisioHyperlink> links in new[] {
                     document.Pages[0].Shapes[0].Hyperlinks, document.Pages[0].Connectors.Single().Hyperlinks,
                     document.Masters.Single().Shape.Hyperlinks
                 }) {
            links.RemoveAt(1);
            links.Insert(0, new VisioHyperlink("https://example.org/new"));
        }
        document.DuplicatePage(document.Pages[0], "Copy");

        foreach (VisioDocument saved in new[] {
                     document, VisioDocument.Load(new MemoryStream(document.ToBytes())),
                     VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value
                 }) {
            foreach (XElement[] section in LegacyRows(saved)) {
                XElement third = section.Single(row => row.Element(Legacy + "Address")!.Value == "https://example.org/third");
                Assert.Equal("7", (string?)third.Attribute("ID"));
                XElement first = section.Single(row => row.Element(Legacy + "Address")!.Value == "https://example.org/first");
                Assert.Equal(indexed ? "9" : null, (string?)first.Attribute("ID"));
                Assert.Equal("Localized first", (string?)first.Attribute("Name"));
                XElement description = first.Element(Legacy + "Description")!;
                Assert.Equal("null", (string?)description.Attribute("V"));
                Assert.Equal("STR", (string?)description.Attribute("Unit"));
                Assert.Equal("description error", (string?)description.Attribute("Err"));
                XElement address = first.Element(Legacy + "Address")!;
                Assert.Equal("STR", (string?)address.Attribute("Unit"));
                Assert.Equal("\"https://example.org/first\"", (string?)address.Attribute("F"));
                Assert.Equal("#VALUE!", (string?)address.Attribute("Err"));
                XElement unnamed = section.Single(row => row.Element(Legacy + "Address")!.Value == "https://example.org/unnamed-source");
                Assert.Equal("12", (string?)unnamed.Attribute("ID"));
                Assert.Null(unnamed.Attribute("NameU"));
            }
            Assert.Equal(5, PackageRows(saved).Count);
            foreach (XElement[] section in PackageRows(saved)) {
                XElement ByAddress(string address) => section.Single(row => row.Elements(Modern + "Cell")
                    .Single(cell => (string?)cell.Attribute("N") == "Address").Attribute("V")!.Value == "https://example.org/" + address);
                XElement added = ByAddress("new");
                Assert.Equal("Row_2", (string?)added.Attribute("N"));
                Assert.Null(added.Attribute("IX"));
                Assert.Equal("Row_3", (string?)ByAddress("third").Attribute("N"));
                Assert.Equal("7", (string?)ByAddress("third").Attribute("IX"));
                Assert.Equal("Row_1", (string?)ByAddress("first").Attribute("N"));
                Assert.Equal(indexed ? "9" : null, (string?)ByAddress("first").Attribute("IX"));
                Assert.Equal("12", (string?)ByAddress("unnamed-source").Attribute("IX"));
                Assert.Null(ByAddress("unnamed-source").Attribute("N"));
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LoadedRootAndChildInstancesReserveTheirOwnInheritedNamesAndKeepNamedOverrides(bool deltasOnly) {
        string xml = "<VisioDocument xmlns='" + Legacy + "'><Masters><Master ID='1' NameU='Group'><Shapes><Shape ID='100' Type='Group'>"
            + "<XForm><Width>4</Width><Height>2</Height></XForm><Hyperlink ID='1' NameU='Row_1'><Address>https://example.org/root</Address></Hyperlink>"
            + "<Shapes><Shape ID='101'><XForm><Width>2</Width><Height>1</Height></XForm>"
            + "<Hyperlink ID='2' NameU='Row_2'><Address>https://example.org/child</Address></Hyperlink></Shape></Shapes></Shape></Shapes></Master></Masters>"
            + "<Pages><Page ID='0' Name='Instances'><Shapes><Shape ID='10' Master='1' Type='Group'><Shapes>"
            + "<Shape ID='11' MasterShape='101'/></Shapes></Shape></Shapes></Page></Pages></VisioDocument>";
        VisioDocument document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml))).Value;
        document.WriteMasterDeltasOnly = deltasOnly;
        VisioShape root = document.Pages[0].Shapes[0];
        VisioShape child = root.Children.Single();
        Assert.Empty(root.Hyperlinks);
        Assert.Empty(child.Hyperlinks);
        Assert.Equal("Row_1", root.MasterShape!.Hyperlinks.Single().RowName);
        Assert.Equal("Row_2", child.MasterShape!.Hyperlinks.Single().RowName);
        root.AddHyperlink("https://example.org/local-root");
        child.AddHyperlink("https://example.org/local-child");
        AssertInstanceRows(document, false);
        root.AddHyperlink("https://example.org/root-override").RowName = "Row_1";
        child.AddHyperlink("https://example.org/child-override").RowName = "Row_2";

        foreach (VisioDocument saved in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) }) {
            AssertInstanceRows(saved, true);
        }
        Assert.Null(root.Hyperlinks[0].RowName);
        Assert.Null(child.Hyperlinks[0].RowName);

        static void AssertInstanceRows(VisioDocument document, bool overrides) {
            foreach (bool legacy in new[] { false, true }) {
                List<XElement[]> sections = legacy ? LegacyRows(document) : PackageRows(document);
                string Name(XElement row) => (string)row.Attribute(legacy ? "NameU" : "N")!;
                string Address(XElement row) => legacy ? row.Element(Legacy + "Address")!.Value :
                    row.Elements(Modern + "Cell").Single(cell => (string?)cell.Attribute("N") == "Address").Attribute("V")!.Value;
                XElement[] root = sections.Single(rows => Address(rows[0]) == "https://example.org/local-root");
                XElement[] child = sections.Single(rows => Address(rows[0]) == "https://example.org/local-child");
                Assert.Equal(overrides ? new[] { "Row_2", "Row_1" } : new[] { "Row_2" }, root.Select(Name));
                Assert.Equal(overrides ? new[] { "Row_1", "Row_2" } : new[] { "Row_1" }, child.Select(Name));
            }
        }
    }

    [Fact]
    public void NewSimpleMasterAndSparseInstanceCopiesKeepDistinctInheritedAndLocalRows() {
        var document = VisioDocument.Create();
        var blueprint = new VisioShape("1", 1, 1, 2, 1, "");
        blueprint.AddHyperlink("https://example.org/master");
        VisioMaster master = document.RegisterMaster("Links", blueprint);
        VisioPage page = document.AddPage("Instance");
        VisioShape instance = page.AddShape("10", master, 3, 3, 2, 1);
        Assert.Empty(instance.Hyperlinks);
        instance.AddHyperlink("https://example.org/instance");
        VisioShape copy = document.DuplicatePage(page, "Copy").Shapes[0];
        Assert.Equal("Row_2", Assert.Single(copy.Hyperlinks).RowName);
        foreach (VisioDocument saved in new[] {
                     document, VisioDocument.Load(new MemoryStream(document.ToBytes())),
                     VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value
                 }) {
            List<XElement[]> native = PackageRows(saved);
            List<XElement[]> legacy = LegacyRows(saved);
            Assert.Equal(3, native.Count);
            Assert.Equal(3, legacy.Count);
            foreach (XElement[] rows in native) {
                XElement row = Assert.Single(rows);
                string address = row.Elements(Modern + "Cell").Single(cell => (string?)cell.Attribute("N") == "Address").Attribute("V")!.Value;
                Assert.Equal(address.EndsWith("/master", StringComparison.Ordinal) ? "Row_1" : "Row_2", (string?)row.Attribute("N"));
            }
            foreach (XElement[] rows in legacy) {
                XElement row = Assert.Single(rows);
                string address = row.Element(Legacy + "Address")!.Value;
                Assert.Equal(address.EndsWith("/master", StringComparison.Ordinal) ? "Row_1" : "Row_2", (string?)row.Attribute("NameU"));
            }
        }
        Assert.Null(blueprint.Hyperlinks.Single().RowName);
        Assert.Null(instance.Hyperlinks.Single().RowName);
    }

    [Fact]
    public void NewCompleteMasterInstancesAndPageCopiesKeepMaterializedSourceIdentityAfterReorderingAndRemoval() {
        var document = VisioDocument.Create();
        var blueprint = new VisioShape("1", 1, 1, 2, 1, "");
        blueprint.AddHyperlink("https://example.org/master");
        blueprint.Children.Add(new VisioShape("2", 1, 1, 1, 1, "Child"));
        VisioMaster master = document.RegisterMaster("Links", blueprint);
        VisioPage page = document.AddPage("Instance");
        VisioShape instance = page.AddShape("10", master, 3, 3, 2, 1);
        Assert.Equal("Row_1", Assert.Single(instance.Hyperlinks).RowName);
        instance.AddHyperlink("https://example.org/instance");
        VisioShape copy = document.DuplicatePage(page, "Copy").Shapes[0];
        Assert.Equal(new[] { "Row_1", "Row_2" }, copy.Hyperlinks.Select(row => row.RowName));
        VisioHyperlink copiedNewLink = copy.Hyperlinks[1];
        copy.Hyperlinks.RemoveAt(1);
        copy.Hyperlinks.Insert(0, copiedNewLink);
        copy.Hyperlinks.RemoveAt(1);
        foreach (VisioDocument saved in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) }) {
            List<XElement[]> sections = PackageRows(saved);
            Assert.Equal(3, sections.Count);
            Assert.Contains(sections, rows => rows.Length == 2 && rows.Select(row => (string?)row.Attribute("N")).SequenceEqual(new[] { "Row_1", "Row_2" }));
            Assert.Contains(sections, rows => rows.Length == 1 && (string?)rows[0].Attribute("N") == "Row_2");
            Assert.Contains(sections, rows => rows.Length == 1 && (string?)rows[0].Attribute("N") == "Row_1");
        }
        Assert.Null(blueprint.Hyperlinks.Single().RowName);
        Assert.Null(instance.Hyperlinks[1].RowName);
    }

    private static void AssertNames(VisioDocument document, string?[] expected, int sectionCount) {
        List<XElement[]> legacy = LegacyRows(document);
        List<XElement[]> native = PackageRows(document);
        Assert.Equal(sectionCount, legacy.Count);
        Assert.Equal(sectionCount, native.Count);
        Assert.All(legacy, rows => Assert.Equal(expected, rows.Select(row => (string?)row.Attribute("NameU"))));
        Assert.All(native, rows => Assert.Equal(expected, rows.Select(row => (string?)row.Attribute("N"))));
    }

    private static List<XElement[]> LegacyRows(VisioDocument document) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value))
        .Descendants(Legacy + "Shape").Select(shape => shape.Elements(Legacy + "Hyperlink").ToArray()).Where(rows => rows.Length > 0).ToList();

    private static List<XElement[]> PackageRows(VisioDocument document) {
        using var stream = new MemoryStream(document.ToBytes());
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var sections = new List<XElement[]>();
        foreach (ZipArchiveEntry entry in archive.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.Ordinal) &&
                     (entry.FullName.StartsWith("visio/pages/", StringComparison.Ordinal) || entry.FullName.StartsWith("visio/masters/", StringComparison.Ordinal)))) {
            using Stream part = entry.Open();
            sections.AddRange(XDocument.Load(part).Descendants(Modern + "Section").Where(section => (string?)section.Attribute("N") == "Hyperlink")
                .Select(section => section.Elements(Modern + "Row").ToArray()));
        }
        return sections;
    }
}
