using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioNativeCellMetadataCloningTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void CopiedGraphsKeepGuardedNativeStateAndRebindFormulas(bool pageCopy, bool namedIds) {
        var document = LoadGraph();
        var source = document.Pages[0];
        var target = pageCopy ? document.DuplicatePage(source, "Copy") : source;
        var copies = pageCopy ? target.Shapes.ToArray() : target.DuplicateShapes(source.Shapes.ToArray(),
            new VisioShapeDuplicationOptions { IdSuffix = namedIds ? "-copy" : null, OffsetX = 0, OffsetY = 0 }).ToArray();
        var root = copies[0];
        root.NameU = "CopiedRoot"; root.Children[0].NameU = "CopiedChild";
        var edge = target.Connectors.Last(); edge.Label = "CopiedEdge";
        target.Shapes.Insert(0, new VisioShape("inserted", 0, 0, 1, 1, ""));
        target.Shapes.Remove(root); target.Shapes.Insert(0, root);
        // The copied state must survive a second graph copy and newly assigned identities.
        var second = target.DuplicateShapes(copies, new VisioShapeDuplicationOptions { IdSuffix = "-again", OffsetX = 0, OffsetY = 0 });
        second[0].NameU = "RepeatedRoot"; second[0].Children[0].NameU = "RepeatedChild";
        target.Connectors.Last().Label = "RepeatedEdge";
        foreach (var candidate in new[] { document, Reopen(document) }) {
            var xml = Export(candidate);
            var original = Shape(xml, "Root");
            AssertState(UserValue(original), "null", "Sheet.002!Width", "STR", "#VALUE!");
            foreach (string prefix in new[] { "Copied", "Repeated" }) {
                var copied = Shape(xml, prefix + "Root");
                var child = Shape(xml, prefix + "Child");
                var connector = xml.Descendants(Legacy + "Shape").Single(s => s.Element(Legacy + "Text")?.Value == prefix + "Edge");
                string childId = (string)child.Attribute("ID")!, rootId = (string)copied.Attribute("ID")!;
                AssertState(UserValue(copied), "null", $"Sheet.{childId}!Width", "STR", "#VALUE!");
                AssertState(copied.Element(Legacy + "Misc")!.Element(Legacy + "Comment")!, "null", $"Sheet.{childId}!Width", "STR", "producer root error");
                AssertState(child.Element(Legacy + "Help")!.Element(Legacy + "HelpTopic")!, "null", null, null, null);
                AssertState(connector.Element(Legacy + "Misc")!.Element(Legacy + "Comment")!, "null", $"Sheet.{rootId}!Width", "STR", "producer edge error");
                AssertState(copied.Element(Legacy + "Para")!.Element(Legacy + "BulletStr")!, "null", "Inh", "STR", null);
            }
            AssertModernCells(candidate.ToBytes());
        }
    }

    [Theory]
    [InlineData("value", false)]
    [InlineData("formula", false)]
    [InlineData("unit", false)]
    [InlineData("value", true)]
    [InlineData("formula", true)]
    [InlineData("unit", true)]
    public void EditsBeforeOrAfterCopySuppressStaleStateWithoutChangingSiblings(string edit, bool afterCopy) {
        var document = LoadGraph();
        var page = document.Pages[0];
        var original = page.Shapes[0];
        if (!afterCopy) Edit(original);
        var clone = page.DuplicateShapes(page.Shapes.ToArray(), new VisioShapeDuplicationOptions { OffsetX = 0, OffsetY = 0 })[0];
        clone.NameU = "EditedCopy";
        if (afterCopy) Edit(clone);
        foreach (var candidate in new[] { document, Reopen(document) }) {
            var xml = Export(candidate);
            var copied = Shape(xml, "EditedCopy");
            var value = UserValue(copied);
            Assert.Null(value.Attribute("V"));
            if (edit == "formula") Assert.Null(value.Attribute("Err"));
            Assert.Null(copied.Element(Legacy + "Para")!.Element(Legacy + "BulletStr")!.Attribute("V"));
            Assert.Equal("null", (string?)copied.Element(Legacy + "Misc")!.Element(Legacy + "Comment")!.Attribute("V"));
            if (afterCopy) Assert.Equal("null", (string?)UserValue(Shape(xml, "Root")).Attribute("V"));
        }
        void Edit(VisioShape shape) {
            var row = shape.FindUserCell("Nullable")!;
            if (edit == "value") row.Value = "edited";
            if (edit == "formula") row.Formula = "Sheet.3!Width";
            if (edit == "unit") row.Unit = "IN";
            var section = shape.GetShapeSheetSections().Single(s => s.Name == "Paragraph");
            section.Rows[0].SetCell("BulletStr", edit == "value" ? "edited" : "", edit == "formula" ? "Sheet.3!Width" : "Inh", edit == "unit" ? "IN" : "STR");
            shape.SetShapeSheetSection(section);
        }
    }

    [Theory]
    [InlineData("same")]
    [InlineData("register")]
    [InlineData("stencil")]
    public void IndependentNxbreMastersCarryTheirThreeMarkersAcrossAllImportAndCopyRoutes(string route) {
        var source = VisioDocument.LoadLegacyXml(Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", "nxbre-ie3.vsx")).Value;
        var (document, master) = Transfer(source, "Atom", route);
        var page = document.AddPage("Instances");
        var instance = page.AddShape("named-instance", master, 1, 1, master.Shape.Width, master.Shape.Height);
        instance.NameU = "Instance";
        var selected = page.DuplicateShapes(new[] { instance }, new VisioShapeDuplicationOptions { IdSuffix = "-copy", OffsetX = 0, OffsetY = 0 })[0];
        selected.NameU = "Selected";
        var duplicatedPage = document.DuplicatePage(page, "Duplicated");
        foreach (var candidate in new[] { document, Reopen(document) }) {
            var xml = Export(candidate);
            var exportedMaster = xml.Descendants(Legacy + "Master").Single(m => (string?)m.Attribute("NameU") == "Atom");
            Assert.Equal(3, NullCount(exportedMaster));
            Assert.Equal(route == "register" ? 0 : 6, xml.Descendants(Legacy + "StyleSheet").Sum(NullCount));
            Assert.Equal(3, NullCount(Shape(xml, "Instance", "Instances")));
            Assert.Equal(3, NullCount(Shape(xml, "Selected", "Instances")));
            Assert.Equal(6, NullCount(xml.Descendants(Legacy + "Page").Single(p => (string?)p.Attribute("Name") == duplicatedPage.Name)));
            AssertModernCells(candidate.ToBytes());
        }
    }

    [Fact]
    public void RegisteredSingleShapeBlueprintKeepsItsOwnNativeState() {
        var source = LoadGraph();
        var blueprint = source.Pages[0].Shapes[0];
        blueprint.Children.Clear(); blueprint.Type = "Shape";
        var document = VisioDocument.Create();
        var master = document.RegisterMaster("Single blueprint", blueprint, "0");
        var page = document.AddPage("Single");
        var instance = page.AddShape("instance", master, 1, 1, 2, 1); instance.NameU = "SingleInstance";
        foreach (var candidate in new[] { document, Reopen(document) }) {
            var xml = Export(candidate);
            AssertState(UserValue(Shape(xml, "SingleInstance")), "null", "Sheet.002!Width", "STR", "#VALUE!");
            Assert.Equal(3, NullCount(xml.Descendants(Legacy + "Master").Single(m => (string?)m.Attribute("NameU") == master.NameU)));
        }
    }

    [Theory]
    [InlineData("value", false)]
    [InlineData("formula", false)]
    [InlineData("unit", false)]
    [InlineData("value", true)]
    [InlineData("formula", true)]
    [InlineData("unit", true)]
    public void ConnectorSectionEditsSuppressStaleNativeState(string edit, bool afterCopy) {
        var document = LoadGraph(); var page = document.Pages[0];
        if (!afterCopy) Edit(page.Connectors[0]);
        page.DuplicateShapes(page.Shapes.ToArray(), new VisioShapeDuplicationOptions { OffsetX = 0, OffsetY = 0 });
        var clone = page.Connectors.Last(); clone.Label = "EditedEdge";
        if (afterCopy) Edit(clone);
        foreach (var candidate in new[] { document, Reopen(document) }) {
            var xml = Export(candidate);
            var copied = xml.Descendants(Legacy + "Shape").Single(s => s.Element(Legacy + "Text")?.Value == "EditedEdge");
            var bullet = copied.Element(Legacy + "Para")!.Element(Legacy + "BulletStr")!;
            Assert.Null(bullet.Attribute("V")); Assert.Null(bullet.Attribute("Err"));
            Assert.Equal("null", (string?)copied.Element(Legacy + "Misc")!.Element(Legacy + "Comment")!.Attribute("V"));
        }
        void Edit(VisioConnector connector) {
            var section = connector.GetShapeSheetSections().Single(s => s.Name == "Paragraph");
            section.Rows[0].SetCell("BulletStr", edit == "value" ? "edited" : "", edit == "formula" ? "Sheet.3!Width" : "Sheet.1!Width", edit == "unit" ? "IN" : "STR");
            connector.SetShapeSheetSection(section);
        }
    }

    [Fact]
    public void CopiedPageSheetUsesItsNewPageIdentity() {
        var document = LoadSheets();
        document.DuplicatePage(document.Pages[0], "Copied sheet");
        foreach (var candidate in new[] { document, Reopen(document) }) {
            var xml = Export(candidate);
            foreach (var page in xml.Descendants(Legacy + "Page"))
                AssertSheet(page.Element(Legacy + "PageSheet")!);
        }
    }

    [Fact]
    public void CopiedPageSheetReferencesFollowTheCopiedGraph() {
        var source = Export(LoadGraph());
        source.Descendants(Legacy + "Page").Single().AddFirst(new XElement(Legacy + "PageSheet",
            new XElement(Legacy + "Misc", new XElement(Legacy + "Comment", new XAttribute("V", "null"), new XAttribute("F", "Sheet.002!Width"), new XAttribute("Unit", "STR"), new XAttribute("Err", "producer sheet error")))));
        var document = Load(source.ToString(SaveOptions.DisableFormatting));
        document.DuplicatePage(document.Pages[0], "Copy");
        foreach (var candidate in new[] { document, Reopen(document) }) {
            var xml = Export(candidate);
            var copied = xml.Descendants(Legacy + "Page").Single(p => (string?)p.Attribute("Name") == "Copy");
            var child = copied.Descendants(Legacy + "Shape").Single(s => (string?)s.Attribute("NameU") == "Child");
            AssertState(copied.Element(Legacy + "PageSheet")!.Element(Legacy + "Misc")!.Element(Legacy + "Comment")!, "null", $"Sheet.{(string)child.Attribute("ID")!}!Width", "STR", "producer sheet error");
        }
    }

    [Theory]
    [InlineData("same")]
    [InlineData("register")]
    [InlineData("stencil")]
    public void MasterPageSheetsCarryOnlyTheirScopedNativeState(string route) {
        var (document, master) = Transfer(LoadSheets(), "Sheet master", route);
        document.AddPage("Use master").AddShape("instance", master, 1, 1, 1, 1);
        foreach (var candidate in new[] { document, Reopen(document) }) {
            var xml = Export(candidate);
            var exported = xml.Descendants(Legacy + "Master").Single(m => (string?)m.Attribute("NameU") == master.NameU);
            AssertSheet(exported.Element(Legacy + "PageSheet")!);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RemovingTheLastMarkerDoesNotChangeSingleBlueprintPreservation(bool removeSection) {
        var source = Load($"<VisioDocument xmlns='{Legacy}'><Pages><Page ID='0'><Shapes><Shape ID='1'><XForm><Width>1</Width><Height>1</Height></XForm>"
            + "<Misc><Comment>literal adjacent cell</Comment></Misc><Para IX='0'><BulletStr V='null' Unit='STR' F='Inh'/></Para></Shape></Shapes></Page></Pages></VisioDocument>");
        var blueprint = source.Pages[0].Shapes[0];
        if (removeSection) blueprint.RemoveShapeSheetSection("Paragraph");
        else {
            var para = blueprint.GetShapeSheetSections().Single(s => s.Name == "Paragraph");
            para.Rows[0].SetCell("BulletStr", "edited", "Inh", "STR"); blueprint.SetShapeSheetSection(para);
        }
        blueprint = Reopen(source).Pages[0].Shapes[0];
        var document = VisioDocument.Create(); var master = document.RegisterMaster("Literal blueprint", blueprint, "0");
        var instance = document.AddPage("Instance").AddShape("instance", master, 1, 1, 1, 1); instance.NameU = "LiteralInstance";
        foreach (var candidate in new[] { document, Reopen(document) }) {
            var xml = Export(candidate);
            var shape = Shape(xml, "LiteralInstance");
            Assert.Equal("literal adjacent cell", shape.Element(Legacy + "Misc")!.Element(Legacy + "Comment")!.Value);
            Assert.Equal(0, NullCount(shape));
            Assert.Equal("literal adjacent cell", xml.Descendants(Legacy + "Master").Single(m => (string?)m.Attribute("NameU") == master.NameU)
                .Element(Legacy + "Shapes")!.Element(Legacy + "Shape")!.Element(Legacy + "Misc")!.Element(Legacy + "Comment")!.Value);
        }
    }

    private static VisioDocument LoadSheets() {
        const string sheet = "<PageSheet><Misc><Comment V='null' Unit='STR' F='Inh' Err='producer sheet error'/></Misc><User ID='0' NameU='Nullable'><Value V='null' Unit='STR' F='Inh' Err='producer sheet error'/></User></PageSheet>";
        return Load($"<VisioDocument xmlns='{Legacy}'><Masters><Master ID='0' NameU='Sheet master'>" + sheet
            + "<Shapes><Shape ID='1'><XForm><Width>1</Width><Height>1</Height></XForm></Shape></Shapes></Master></Masters><Pages><Page ID='0' Name='Source'>" + sheet + "</Page></Pages></VisioDocument>", VisioPackageType.Template);
    }

    [Theory]
    [InlineData("same")]
    [InlineData("register")]
    [InlineData("stencil")]
    public void MasterInstantiationRebindsNullAndProducerErrorSnapshotsMechanically(string route) {
        var source = LoadGraph(withMaster: true);
        var (document, master) = Transfer(source, "Null master", route);
        var page = document.AddPage("Master instances");
        var instance = page.AddShape("named-instance", master, 1, 1, 2, 1);
        instance.NameU = "Instance"; instance.Children[0].NameU = "InstanceChild";
        foreach (var candidate in new[] { document, Reopen(document) }) {
            var xml = Export(candidate);
            var root = Shape(xml, "Instance"); var child = Shape(xml, "InstanceChild");
            string formula = $"Sheet.{(string)child.Attribute("ID")!}!Width";
            AssertState(UserValue(root), "null", formula, "STR", "#VALUE!");
            AssertState(root.Element(Legacy + "Misc")!.Element(Legacy + "Comment")!, "null", formula, "STR", "producer root error");
            var masterXml = xml.Descendants(Legacy + "Master").Single(m => (string?)m.Attribute("NameU") == master.NameU);
            AssertState(UserValue(masterXml.Descendants(Legacy + "Shape").First()), "null", "Sheet.002!Width", "STR", "#VALUE!");
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StencilStyleMetadataTransfersOnlyToActuallyImportedCells(bool collision) {
        string style = "<StyleSheets><StyleSheet ID='10' NameU='Imported'><Misc><Comment {0}/></Misc></StyleSheet></StyleSheets>";
        var source = Load($"<VisioDocument xmlns='{Legacy}'>" + string.Format(style, "V='null'") + "<Masters><Master ID='0' NameU='Tiny'><Shapes><Shape ID='1'><XForm><Width>1</Width><Height>1</Height></XForm></Shape></Shapes></Master></Masters></VisioDocument>", VisioPackageType.Stencil);
        var destination = collision ? Load($"<VisioDocument xmlns='{Legacy}'>" + string.Format(style, "") + "</VisioDocument>") : VisioDocument.Create();
        string path = PackagePath();
        try {
            File.WriteAllBytes(path, source.ToBytes()); destination.ImportStencilMastersAndGet(path, new[] { "Tiny" });
            foreach (var candidate in new[] { destination, Reopen(destination) }) {
                var comment = Export(candidate).Descendants(Legacy + "StyleSheet").Single(s => (string?)s.Attribute("ID") == "10")
                    .Element(Legacy + "Misc")!.Element(Legacy + "Comment")!;
                Assert.Equal(collision ? null : "null", (string?)comment.Attribute("V"));
            }
        } finally { File.Delete(path); }
    }

    [Fact]
    public void ImportedFontIdentityRebindingPreservesGuardedProducerErrors() {
        const string character = "<Char IX='0'><Font F='GUARD(5)' Err='producer font error'>5</Font></Char>";
        var source = Load($"<VisioDocument xmlns='{Legacy}'><FaceNames><FaceName ID='5' Name='Source family'/></FaceNames>"
            + "<StyleSheets><StyleSheet ID='10' NameU='Source style'>" + character + "</StyleSheet></StyleSheets>"
            + "<Masters><Master ID='0' NameU='Font master'><Shapes><Shape ID='1'><XForm><Width>1</Width><Height>1</Height></XForm>" + character + "</Shape></Shapes></Master></Masters></VisioDocument>", VisioPackageType.Stencil);
        var destination = Load($"<VisioDocument xmlns='{Legacy}'><FaceNames><FaceName ID='5' Name='Destination family'/></FaceNames></VisioDocument>");
        string path = PackagePath();
        try {
            File.WriteAllBytes(path, source.ToBytes());
            var master = destination.ImportStencilMastersAndGet(path, new[] { "Font master" }).Single();
            destination.AddPage("Font instance").AddShape("font-instance", master, 1, 1, 1, 1);
            foreach (var candidate in new[] { destination, Reopen(destination) }) {
                var xml = Export(candidate);
                string mapped = (string)xml.Descendants(Legacy + "FaceName").Single(f => (string?)f.Attribute("Name") == "Source family").Attribute("ID")!;
                Assert.NotEqual("5", mapped);
                var definition = xml.Descendants(Legacy + "Master").Single(m => (string?)m.Attribute("NameU") == "Font master");
                var style = xml.Descendants(Legacy + "StyleSheet").Single(s => (string?)s.Attribute("ID") == "10");
                foreach (var font in definition.Descendants(Legacy + "Font").Concat(style.Descendants(Legacy + "Font"))) {
                    Assert.Equal(mapped, font.Value); Assert.Equal("GUARD(" + mapped + ")", (string?)font.Attribute("F"));
                    Assert.Equal("producer font error", (string?)font.Attribute("Err"));
                }
            }
            // A caller changes the family after import; the old producer error is no longer applicable.
            master.Shape.TextStyle!.FontFamily = "Destination family";
            var edited = Export(Reopen(destination)).Descendants(Legacy + "Master").Single(m => (string?)m.Attribute("NameU") == "Font master");
            Assert.All(edited.Descendants(Legacy + "Font"), font => Assert.Null(font.Attribute("Err")));
        } finally { File.Delete(path); }
    }

    private static (VisioDocument, VisioMaster) Transfer(VisioDocument source, string name, string route) {
        if (route == "same") return (source, source.GetMaster(name));
        var destination = VisioDocument.Create();
        if (route == "register") return (destination, destination.RegisterMaster(source.GetMaster(name)));
        string path = PackagePath();
        try { File.WriteAllBytes(path, source.ToBytes()); return (destination, destination.ImportStencilMastersAndGet(path, new[] { name }).Single()); }
        finally { File.Delete(path); }
    }

    private static string PackagePath() => Path.Combine(AppContext.BaseDirectory, "native-cell-" + Guid.NewGuid().ToString("N") + ".vssx");
    private static VisioDocument LoadGraph(bool withMaster = false) {
        const string root = "<Shape ID='1' NameU='Root' Type='Group'><XForm><Width>2</Width><Height>1</Height></XForm>"
            + "<Misc><Comment V='null' Unit='STR' F='Sheet.002!Width' Err='producer root error'/></Misc>"
            + "<Para IX='0'><BulletStr V='null' Unit='STR' F='Inh'/></Para>"
            + "<User ID='0' NameU='Nullable'><Value V='null' Unit='STR' F='Sheet.002!Width' Err='#VALUE!'/></User>"
            + "<Shapes><Shape ID='002' NameU='Child'><XForm><Width>1</Width><Height>1</Height></XForm><Help><HelpTopic V='null'/></Help></Shape></Shapes></Shape>";
        string container = withMaster ? "<Masters><Master ID='0' NameU='Null master'><Shapes>" + root + "</Shapes></Master></Masters>"
            : "<Pages><Page ID='0' Name='Source'><Shapes>" + root
            + "<Shape ID='3' NameU='Other'><XForm><Width>1</Width><Height>1</Height></XForm></Shape>"
            + "<Shape ID='4'><XForm1D><BeginX>0</BeginX><BeginY>0</BeginY><EndX>2</EndX><EndY>0</EndY></XForm1D>"
            + "<Misc><Comment V='null' Unit='STR' F='Sheet.1!Width' Err='producer edge error'/></Misc>"
            + "<Para IX='0'><BulletStr V='null' Unit='STR' F='Sheet.1!Width' Err='producer paragraph error'/></Para><Text>Edge</Text></Shape>"
            + "</Shapes><Connects><Connect FromSheet='4' FromCell='BeginX' ToSheet='1' ToCell='PinX'/><Connect FromSheet='4' FromCell='EndX' ToSheet='3' ToCell='PinX'/></Connects></Page></Pages>";
        return Load($"<VisioDocument xmlns='{Legacy}'>" + container + "</VisioDocument>", withMaster ? VisioPackageType.Stencil : VisioPackageType.Drawing);
    }

    private static VisioDocument Load(string xml, VisioPackageType type = VisioPackageType.Drawing) => VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml)), type).Value;
    private static VisioDocument Reopen(VisioDocument document) => VisioDocument.Load(new MemoryStream(document.ToBytes()));
    private static XDocument Export(VisioDocument document) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
    private static XElement Shape(XDocument xml, string name, string? page = null) => (page == null ? xml.Descendants(Legacy + "Page") : xml.Descendants(Legacy + "Page").Where(p => (string?)p.Attribute("Name") == page))
        .Descendants(Legacy + "Shape").Single(s => (string?)s.Attribute("NameU") == name);
    private static XElement UserValue(XElement shape) => shape.Elements(Legacy + "User").Single(r => (string?)r.Attribute("NameU") == "Nullable").Element(Legacy + "Value")!;
    private static void AssertSheet(XElement sheet) {
        AssertState(UserValue(sheet), "null", "Inh", "STR", "producer sheet error");
        AssertState(sheet.Element(Legacy + "Misc")!.Element(Legacy + "Comment")!, "null", "Inh", "STR", "producer sheet error");
    }
    private static int NullCount(XElement root) => root.DescendantsAndSelf().Count(c => (string?)c.Attribute("V") == "null");
    private static void AssertState(XElement cell, string? marker, string? formula, string? unit, string? error) {
        Assert.Equal("", cell.Value); Assert.Equal(marker, (string?)cell.Attribute("V")); Assert.Equal(formula, (string?)cell.Attribute("F"));
        Assert.Equal(unit, (string?)cell.Attribute("Unit")); Assert.Equal(error, (string?)cell.Attribute("Err"));
    }
    private static void AssertModernCells(byte[] bytes) {
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        foreach (var entry in archive.Entries.Where(e => e.FullName.StartsWith("visio/", StringComparison.Ordinal) && e.FullName.EndsWith(".xml", StringComparison.Ordinal))) {
            using var stream = entry.Open(); var xml = XDocument.Load(stream);
            foreach (var cell in xml.Descendants(Modern + "Cell"))
                Assert.All(cell.Attributes().Where(a => !a.IsNamespaceDeclaration), a => { Assert.Equal(XNamespace.None, a.Name.Namespace); Assert.Contains(a.Name.LocalName, new[] { "N", "V", "U", "E", "F" }); });
        }
    }
}
