using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class VisioMasterReplacementTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";

    [Theory]
    [InlineData(false, true, "vdx")]
    [InlineData(false, false, "vdx")]
    [InlineData(false, true, "vsdx")]
    [InlineData(false, false, "vsdx")]
    [InlineData(true, true, "vdx")]
    [InlineData(true, false, "vdx")]
    [InlineData(true, true, "vsdx")]
    [InlineData(true, false, "vsdx")]
    public void ReplacementRebindsReorderedChildrenAndKeepsLiveIdentityAndGlue(bool resize, bool deltas, string format) {
        var definitions = VisioDocument.Create(VisioPackageType.Stencil);
        definitions.RegisterMaster("Original", Blueprint("10", "11", "12", "13", 4, 2, "Rectangle", "Triangle"));
        definitions.RegisterMaster("Replacement", Blueprint("20", "30", "31", "32", 8, 4, "Ellipse", "Diamond"));
        definitions = VisioDocument.Load(new MemoryStream(definitions.ToBytes()));
        var document = VisioDocument.Create(); document.WriteMasterDeltasOnly = deltas;
        VisioMaster original = document.RegisterMaster(definitions.GetMaster("Original"));
        VisioMaster replacement = document.RegisterMaster(definitions.GetMaster("Replacement"));
        VisioPage page = document.AddPage("Page", 12, 8);
        VisioShape root = page.AddShape("root", original, 5, 4, 4, 2);
        VisioShape child = root.Children[0], nested = root.Children[1], grandchild = nested.Children[0];
        root.Children.Remove(nested); root.Children.Insert(0, nested);
        child.Text = "Retained label"; child.SetShapeData("Owner", "Operations");
        child.SetUserCell("Approval", "complete"); child.LayerNames.Add("Review");
        child.AddHyperlink("https://example.org/item"); child.Protection.Deletion();
        child.FillColor = OfficeColor.LightBlue;
        child.TextStyle = new VisioTextStyle { Size = 12, TextWidth = .5, TextPinX = .6 };
        VisioTextStyle style = child.TextStyle;
        var point = new VisioConnectionPoint(.8, .25, 1, 0); child.ConnectionPoints.Add(point);
        VisioShape target = page.AddRectangle(10, 4, 1, 1);
        VisioConnector connector = page.AddConnector(child, target, ConnectorKind.Straight);
        connector.FromConnectionPoint = point;
        string originalSvg = page.ToSvg(RenderOptions());
        string masterBefore = MasterDefinition(definitions.GetMaster("Original"));

        // Selecting a root and its descendants applies one tree plan.
        page.SelectShapes(shape => shape == root || shape == child).ReplaceMaster(replacement, resize);
        Assert.Same(nested, root.Children[0]); Assert.Same(child, root.Children[1]);
        Assert.Same(grandchild, nested.Children[0]); Assert.Same(child, connector.From);
        Assert.Same(point, connector.FromConnectionPoint); Assert.Contains(point, child.ConnectionPoints);
        Assert.Same(style, child.TextStyle);
        Assert.Same(replacement.Shape.Children[0], child.MasterShape);
        Assert.Same(replacement.Shape.Children[1].Children[0], grandchild.MasterShape);
        Assert.Equal("30", child.MasterShapeId); Assert.Equal("32", grandchild.MasterShapeId);
        Assert.Equal(masterBefore, MasterDefinition(definitions.GetMaster("Original")));
        double factor = resize ? 2 : 1;
        Assert.Equal(4 * factor, root.Width); Assert.Equal(2 * factor, root.Height);
        Assert.Equal(1.5 * factor, child.Width); Assert.Equal(.8 * factor, point.X);
        Assert.Equal(.5 * factor, style.TextWidth); Assert.Equal(12, style.Size);
        Assert.NotEqual(originalSvg, page.ToSvg(RenderOptions()));

        foreach (VisioDocument candidate in new[] { document, Reopen(document, format) }) {
            VisioShape saved = candidate.Pages[0].FindShapeById(child.Id)!;
            Assert.Equal("Replacement", saved.MasterNameU); Assert.Equal("Ellipse", saved.MasterShape!.NameU);
            Assert.Equal("30", saved.MasterShapeId); Assert.Equal("Retained label", saved.Text);
            Assert.Equal("Operations", saved.GetShapeDataValue("Owner"));
            Assert.Equal("complete", saved.GetUserCellValue("Approval"));
            Assert.Contains("Review", saved.LayerNames); Assert.True(saved.Protection.LockDelete);
            Assert.Equal("https://example.org/item", Assert.Single(saved.Hyperlinks).Address);
            Assert.Equal(1.5 * factor, saved.Width, 8); Assert.Equal(.5 * factor, saved.TextStyle!.TextWidth);
            VisioConnector glue = Assert.Single(candidate.Pages[0].Connectors);
            Assert.Same(saved, glue.From); Assert.Equal(.8 * factor, glue.FromConnectionPoint!.X, 8);
            Assert.Equal(connector.StartPoint.X, glue.StartPoint.X, 8); Assert.Equal(connector.StartPoint.Y, glue.StartPoint.Y, 8);
            XElement rendered = XDocument.Parse(candidate.Pages[0].ToSvg(RenderOptions())).Descendants(Svg + "g")
                .Single(g => (string?)g.Attribute("data-visio-shape-id") == child.Id);
            Assert.DoesNotContain(rendered.Elements(Svg + "rect"), rect => (string?)rect.Attribute("data-officeimo-preserved-geometry") == "true");
            Assert.NotEmpty(rendered.Elements(Svg + "path"));
        }
    }

    [Theory]
    [InlineData("name")]
    [InlineData("master")]
    [InlineData("stencil")]
    public void IncompatibleSelectionDoesNotMutateEarlierShapesOrRegisterReplacement(string overload) {
        var document = VisioDocument.Create(); var page = document.AddPage("Page");
        VisioShape first = page.AddRectangle(1, 1, 1, 1, "first");
        var group = new VisioShape("group", 4, 4, 2, 2, "group");
        group.Children.Add(new VisioShape("child", 1, 1, 1, 1, "child")); page.Shapes.Add(group);
        byte[] before = document.ToLegacyXmlResult().Value;
        var masters = document.Masters.ToArray();
        var replacement = new VisioMaster("new", "Ellipse", new VisioShape("1", .5, .5, 1, 1, "new"));
        VisioShapeSelection selected = page.SelectShapes(shape => shape == first || shape == group);
        Assert.Throws<NotSupportedException>(() => {
            if (overload == "name") selected.ReplaceMaster("Ellipse");
            else if (overload == "master") selected.ReplaceMaster(replacement);
            else selected.ReplaceMaster(new VisioStencilShape("oval", "Oval", "Ellipse", "Shapes", 1, 1));
        });
        Assert.Equal(before, document.ToLegacyXmlResult().Value);
        Assert.Equal(masters, document.Masters.ToArray()); Assert.False(document.TryGetMaster("Ellipse", out _));
        Assert.Equal("Rectangle", first.MasterNameU); Assert.Same(group.Children[0], page.FindShapeById("child"));
        Assert.Null(replacement.Shape.PersistedId);
    }

    [Fact]
    public void LaterUnsupportedFrameLeavesEarlierSelectedGroupsUnchanged() {
        var document = VisioDocument.Create();
        VisioMaster original = document.RegisterMaster("Original", Blueprint("1", "2", "3", "4", 4, 2, "Rectangle", "Triangle"));
        var replacement = new VisioMaster("replacement", "Replacement", Blueprint("5", "6", "7", "8", 8, 2, "Ellipse", "Diamond"));
        var page = document.AddPage("Page");
        VisioShape first = page.AddShape("first", original, 3, 3, 4, 2);
        VisioShape second = page.AddShape("second", original, 3, 6, 4, 2);
        second.Children[1].Angle = .3;
        byte[] before = document.ToLegacyXmlResult().Value;
        Assert.Throws<NotSupportedException>(() => page.SelectByMaster("Original").ReplaceMaster(replacement, resizeToMaster: true));
        Assert.Equal(before, document.ToLegacyXmlResult().Value); Assert.Same(original, first.Master);
        Assert.False(document.TryGetMaster("Replacement", out _)); Assert.Null(replacement.Shape.PersistedId);
    }

    [Theory]
    [InlineData("vdx")]
    [InlineData("vsdx")]
    public void RetainedInheritedTextAndFormattingBecomeLocal(string format) {
        string source = "<VisioDocument xmlns='" + Legacy + "'><Fonts><FontEntry ID='0' Name='Arial'/></Fonts><Masters>" +
            "<Master ID='1' NameU='Rectangle'><Shapes><Shape ID='1'><XForm><Width>2</Width><Height>1</Height></XForm><Char IX='7'><Font>0</Font><Size>0.1666666666666667</Size><Style>17</Style><LangID>1033</LangID></Char><Para IX='9'><HorzAlign>1</HorzAlign><SpLine>1</SpLine></Para><Text><cp IX='7'/><pp IX='9'/>Retained text</Text></Shape></Shapes></Master>" +
            "<Master ID='2' NameU='Ellipse'><Shapes><Shape ID='4'><XForm><Width>2</Width><Height>1</Height></XForm><Char IX='7'><Size>0.5</Size><Style>0</Style></Char><Para IX='9'><HorzAlign>2</HorzAlign></Para><Text>Replacement label</Text></Shape></Shapes></Master>" +
            "</Masters><Pages><Page ID='0'><Shapes><Shape ID='10' Master='1'><XForm><PinX>2</PinX><PinY>2</PinY></XForm></Shape></Shapes></Page></Pages></VisioDocument>";
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
        VisioShape shape = document.Pages[0].Shapes[0]; Assert.True(shape.HasInheritedText);
        document.Pages[0].ReplaceMaster(shape, document.GetMaster("Ellipse"));
        Assert.False(shape.HasInheritedText);
        foreach (VisioDocument candidate in new[] { document, Reopen(document, format) }) {
            VisioShape retained = candidate.Pages[0].Shapes[0];
            Assert.Equal("Retained text", retained.Text); Assert.Equal(12, retained.TextStyle!.Size!.Value, 8);
            Assert.Equal(VisioTextHorizontalAlignment.Center, retained.TextStyle.HorizontalAlignment);
            XElement pageShape = XDocument.Load(new MemoryStream(candidate.ToLegacyXmlResult().Value)).Descendants(Legacy + "Page").Single().Descendants(Legacy + "Shape").Single();
            Assert.Equal("7", Assert.Single(pageShape.Elements(Legacy + "Char")).Attribute("IX")!.Value);
            Assert.Equal("9", Assert.Single(pageShape.Elements(Legacy + "Para")).Attribute("IX")!.Value);
            Assert.Equal("17", pageShape.Element(Legacy + "Char")!.Element(Legacy + "Style")!.Value);
            Assert.Equal("1033", pageShape.Element(Legacy + "Char")!.Element(Legacy + "LangID")!.Value);
            Assert.Equal("Retained text", pageShape.Element(Legacy + "Text")!.Value);
        }
    }

    private static VisioShape Blueprint(string rootId, string childId, string nestedId, string grandchildId, double width, double height, string childKind, string grandchildKind) {
        var root = new VisioShape(rootId, width / 2, height / 2, width, height, "") { FillPattern = 0, LinePattern = 0 };
        root.Children.Add(new VisioShape(childId, width / 4, height / 2, width * .375, height / 2, "") { NameU = childKind });
        var nested = new VisioShape(nestedId, width * .75, height / 2, width / 2, height, "") { FillPattern = 0, LinePattern = 0 };
        nested.Children.Add(new VisioShape(grandchildId, width / 4, height / 2, width * .375, height / 2, "") { NameU = grandchildKind });
        root.Children.Add(nested); return root;
    }
    private static string MasterDefinition(VisioMaster master) {
        var owner = VisioDocument.Create(VisioPackageType.Stencil); owner.RegisterMaster(master);
        return XDocument.Load(new MemoryStream(owner.ToLegacyXmlResult().Value)).Descendants(Legacy + "Master").Single().ToString();
    }
    private static VisioSvgSaveOptions RenderOptions() => new() { RenderText = false, RenderStencilArtwork = false };
    private static VisioDocument Reopen(VisioDocument document, string format) => format == "vdx"
        ? VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value : VisioDocument.Load(new MemoryStream(document.ToBytes()));
}
