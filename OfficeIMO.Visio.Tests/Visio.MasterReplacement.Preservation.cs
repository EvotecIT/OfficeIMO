using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class VisioMasterReplacementTests {
    [Theory]
    [InlineData("vdx")]
    [InlineData("vsdx")]
    public void ProducerGroupReplacementKeepsChildRowsAndRebindsDefinitions(string format) {
        var source = VisioDocument.LoadLegacyXml(Fixture("nxbre-ie3.vsx")).Value;
        VisioMaster old = source.GetMaster("Atom"), replacement = source.GetMaster("Negative Atom");
        var document = VisioDocument.Create(); var page = document.AddPage("Page", 8, 8);
        VisioShape root = page.AddShape("group", old, 4, 4, old.Shape.Width, old.Shape.Height);
        VisioShape child = root.Children[0]; child.Text = "Edited relation";
        child.SetUserCell("Approval", "reviewed"); child.TextStyle!.Size = 12;
        string masterBefore = MasterDefinition(old);
        page.ReplaceMaster(root, replacement);
        Assert.Equal(masterBefore, MasterDefinition(old));
        Assert.Same(child, root.Children[0]); Assert.Same(replacement.Shape.Children[0], child.MasterShape);
        foreach (VisioDocument candidate in new[] { document, Reopen(document, format) }) {
            VisioShape saved = candidate.Pages[0].FindShapeById(child.Id)!;
            Assert.Equal("Negative Atom", saved.MasterNameU);
            Assert.Equal("Edited relation", saved.Text); Assert.Equal(12, saved.TextStyle!.Size);
            Assert.Equal("reviewed", saved.GetUserCellValue("Approval"));
            Assert.Equal(child.MasterShapeId, saved.MasterShapeId);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PackageStencilReplacementImportsOnlyAfterEverySelectedShapePasses(bool incompatible) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-master-replacement-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "source.vssx");
            var source = VisioDocument.LoadLegacyXml(Fixture("nxbre-ie3.vsx")).Value;
            source.Save(path);
            VisioMaster prior = source.GetMaster("Atom"), next = source.GetMaster("Negative Atom");
            var document = VisioDocument.Create(); var page = document.AddPage("Page");
            VisioShape first = page.AddShape("first", prior, 2, 2, prior.Shape.Width, prior.Shape.Height);
            VisioShape second = page.AddShape("second", prior, 5, 5, prior.Shape.Width, prior.Shape.Height);
            if (incompatible) second.Children.Add(new VisioShape("extra", .5, .5, .25, .25, "extra"));
            var stencil = new VisioStencilShape("negative", "Negative Atom", "Negative Atom", "Rules", next.Shape.Width, next.Shape.Height,
                null, null, null, null, VisioMeasurementUnit.Inches, path);
            byte[] before = document.ToLegacyXmlResult().Value;
            var selected = page.SelectShapes(shape => shape == first || shape == second);
            if (incompatible) {
                Assert.Throws<NotSupportedException>(() => selected.ReplaceMaster(stencil));
                Assert.Equal(before, document.ToLegacyXmlResult().Value);
                Assert.False(document.TryGetMaster("Negative Atom", out _));
                Assert.Equal("Atom", first.MasterNameU);
            } else {
                selected.ReplaceMaster(stencil);
                Assert.True(document.GetMaster("Negative Atom").IsPackageBacked);
                Assert.Equal("negative", first.GetUserCellValue(VisioSemanticUserCells.StencilId));
                foreach (VisioDocument candidate in new[] { Reopen(document, "vdx"), Reopen(document, "vsdx") }) {
                    VisioShape saved = candidate.Pages[0].FindShapeById(first.Id)!;
                    Assert.Equal("Negative Atom", saved.MasterNameU);
                    Assert.All(saved.Children, child => Assert.Equal("Negative Atom", child.MasterNameU));
                    Assert.Equal(first.Children[0].TextStyle!.FontFamily, saved.Children[0].TextStyle!.FontFamily);
                }
            }
        } finally { Directory.Delete(directory, recursive: true); }
    }

    [Theory]
    [InlineData("vdx")]
    [InlineData("vsdx")]
    public void ForeignArtworkCanReplaceAndBeReplacedWithoutStaleLocalPayload(string format) {
        var producer = VisioDocument.LoadLegacyXml(Fixture("angularjs-simple-scope.vdx")).Value;
        VisioShape image = producer.Pages[0].AllShapes().Single(shape => shape.Type == "Foreign");
        var document = VisioDocument.Create(); var page = document.AddPage("Page", 8, 8);
        var replacement = new VisioMaster("image", "Embedded image", image);
        VisioShape live = page.AddRectangle(4, 4, image.Width, image.Height, "Kept text");
        live.FillColor = OfficeColor.LightBlue;
        page.ReplaceMaster(live, replacement);
        foreach (VisioDocument candidate in new[] { document, Reopen(document, format) }) {
            XElement svg = XDocument.Parse(candidate.Pages[0].ToSvg(RenderOptions())).Root!;
            Assert.Single(svg.Descendants(Svg + "image"));
            VisioShape saved = candidate.Pages[0].FindShapeById(live.Id)!;
            Assert.Equal("Foreign", saved.Type); Assert.Equal("Kept text", saved.Text);
        }
        // A loaded local payload must not override the replacement master artwork.
        document = Reopen(document, format); page = document.Pages[0]; live = page.FindShapeById(live.Id)!;
        page.ReplaceMaster(live, "Rectangle");
        foreach (VisioDocument candidate in new[] { document, Reopen(document, format) }) {
            Assert.Empty(XDocument.Parse(candidate.Pages[0].ToSvg(RenderOptions())).Descendants(Svg + "image"));
            VisioShape saved = candidate.Pages[0].FindShapeById(live.Id)!;
            Assert.Equal("Rectangle", saved.MasterNameU); Assert.Equal("Kept text", saved.Text);
            Assert.Equal(OfficeColor.LightBlue, saved.FillColor);
        }
    }

    [Fact]
    public void ReplacementRejectsABlueprintThatIsItselfAnEditedInstance() {
        var document = VisioDocument.Create(); var page = document.AddPage("Page");
        VisioShape shape = page.AddRectangle(2, 2, 2, 1, "live");
        byte[] before = document.ToLegacyXmlResult().Value;
        var replacement = new VisioMaster("shared", "Shared", shape);
        Assert.Throws<NotSupportedException>(() => page.ReplaceMaster(shape, replacement));
        Assert.Equal(before, document.ToLegacyXmlResult().Value); Assert.False(document.TryGetMaster("Shared", out _));
    }

    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", name);

    [Theory]
    [InlineData(false, "vdx")]
    [InlineData(false, "vsdx")]
    [InlineData(true, "vdx")]
    [InlineData(true, "vsdx")]
    public void MaterializedInheritedTextRebindsMasterSheetsAndKeepsLocalSheets(bool local, string format) {
        string rows = local ? "<Char IX='0'><Size F='Sheet.2!Width/6'>1</Size></Char><Para IX='0'><IndLeft F='Sheet.2!Width/40'>0.15</IndLeft></Para>" : "";
        string xml = "<VisioDocument xmlns='" + Legacy + "'><Masters>" +
            "<Master ID='1' NameU='Old'><Shapes><Shape ID='1' Type='Group'><XForm><Width>4</Width><Height>2</Height></XForm><Shapes>" +
            "<Shape ID='2' NameU='Rectangle'><XForm><Width>1</Width><Height>1</Height></XForm><Char IX='0'><Size F='Sheet.2!Width/6'>0.1666666666666667</Size></Char><Para IX='0'><IndLeft F='Sheet.1!Width/40'>0.1</IndLeft></Para><Text>Kept</Text></Shape>" +
            "</Shapes></Shape></Shapes></Master>" +
            "<Master ID='3' NameU='New'><Shapes><Shape ID='3' Type='Group'><XForm><Width>4</Width><Height>2</Height></XForm><Shapes>" +
            "<Shape ID='4' NameU='Ellipse'><XForm><Width>1</Width><Height>1</Height></XForm><Char IX='0'><Size>0.5</Size></Char></Shape>" +
            "</Shapes></Shape></Shapes></Master></Masters>" +
            "<Pages><Page ID='0'><Shapes><Shape ID='10' Master='1' Type='Group'><XForm><PinX>2</PinX><PinY>2</PinY></XForm><Shapes>" +
            "<Shape ID='11' MasterShape='2'>" + rows + "<User IX='0' NameU='Local'><Value F='Sheet.2!Width'>6</Value></User></Shape></Shapes></Shape>" +
            "<Shape ID='2'><XForm><Width>6</Width><Height>1</Height></XForm></Shape></Shapes></Page></Pages></VisioDocument>";
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml))).Value;
        document.Pages[0].ReplaceMaster(document.Pages[0].FindShapeById("10")!, document.GetMaster("New"));
        foreach (VisioDocument candidate in new[] { document, Reopen(document, format) }) {
            VisioShape child = candidate.Pages[0].FindShapeById("11")!;
            Assert.Equal(local ? 72 : 12, child.TextStyle!.Size!.Value, 8);
            XElement saved = XDocument.Load(new MemoryStream(candidate.ToLegacyXmlResult().Value)).Descendants(Legacy + "Page").Single()
                .Descendants(Legacy + "Shape").Single(shape => (string?)shape.Attribute("ID") == "11");
            Assert.Equal(local ? "Sheet.2!Width/6" : "Sheet.11!Width/6", (string?)saved.Element(Legacy + "Char")!.Element(Legacy + "Size")!.Attribute("F"));
            Assert.Equal(local ? "Sheet.2!Width/40" : "Sheet.10!Width/40", (string?)saved.Element(Legacy + "Para")!.Element(Legacy + "IndLeft")!.Attribute("F"));
            Assert.Equal("Sheet.2!Width", (string?)saved.Element(Legacy + "User")!.Element(Legacy + "Value")!.Attribute("F"));
        }
    }

    [Theory]
    [InlineData(false, "vdx")]
    [InlineData(false, "vsdx")]
    [InlineData(true, "vdx")]
    [InlineData(true, "vsdx")]
    public void ReplacingALeafMaterializesTheNewGroupWithoutReplacingItsRoot(bool resize, string format) {
        var document = VisioDocument.Create(); var page = document.AddPage("Page", 10, 10);
        VisioShape live = page.AddRectangle(4, 4, 2, 1, "Retained root");
        var replacement = new VisioMaster("composite", "Composite", Blueprint("20", "21", "22", "23", 4, 2, "Ellipse", "Triangle"));
        var peer = page.AddRectangle(8, 4, 1, 1);
        VisioConnector connector = page.AddConnector(live, peer, ConnectorKind.Straight);
        page.ReplaceMaster(live, replacement, resize);
        Assert.Same(live, page.Shapes[0]); Assert.Same(live, connector.From);
        Assert.Equal(2, live.Children.Count); Assert.Single(live.Children[1].Children);
        double factor = resize ? 1 : .5;
        Assert.Equal(1.5 * factor, live.Children[0].Width);
        Assert.Equal(1 * factor, live.Children[0].PinX);
        foreach (VisioDocument candidate in new[] { document, Reopen(document, format) }) {
            VisioShape saved = candidate.Pages[0].FindShapeById(live.Id)!;
            Assert.Equal("Retained root", saved.Text); Assert.Equal("Group", saved.Type);
            Assert.Equal(2, saved.Children.Count); Assert.Equal("Composite", saved.Children[0].MasterNameU);
            Assert.Equal("Ellipse", saved.Children[0].MasterShape!.NameU);
            Assert.Equal("Triangle", saved.Children[1].Children[0].MasterShape!.NameU);
            Assert.Equal(1.5 * factor, saved.Children[0].Width, 8);
            Assert.Same(saved, Assert.Single(candidate.Pages[0].Connectors).From);
        }
    }
}
