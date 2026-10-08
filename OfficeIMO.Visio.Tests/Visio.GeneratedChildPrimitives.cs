using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioGeneratedChildPrimitivesTests {
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";

    [Theory]
    [InlineData(true, "vsdx")]
    [InlineData(false, "vsdx")]
    [InlineData(true, "vdx")]
    [InlineData(false, "vdx")]
    public void GeneratedChildrenKeepTheirPrimitivesBeforeAndAfterSaving(bool deltasOnly, string format) {
        var document = VisioDocument.Create();
        document.WriteMasterDeltasOnly = deltasOnly;
        var master = document.RegisterMaster("Composite", Blueprint());
        VisioPage page = document.AddPage("Page", 8, 4);
        VisioShape instance = page.AddShape("instance", master, 4, 2, 6, 2);
        // Instance labels cannot replace the geometry identity of their linked blueprints.
        instance.Children[0].NameU = "Custom oval label";
        instance.Children[1].NameU = "Custom decision label";
        instance.Children[2].Children[0].NameU = "Custom triangle label";
        AssertSilhouettes(page);

        VisioDocument reopened = Reopen(document, format);
        AssertSilhouettes(reopened.Pages[0]);
        VisioShape reopenedInstance = reopened.Pages[0].FindShapeById("instance")!;
        Assert.Equal("Composite", reopenedInstance.MasterNameU);
        Assert.Equal("Ellipse", reopenedInstance.Children[0].MasterShape!.NameU);
        Assert.Equal("Diamond", reopenedInstance.Children[1].MasterShape!.NameU);
        Assert.Equal("Triangle", reopenedInstance.Children[2].Children[0].MasterShape!.NameU);
    }

    [Fact]
    public void RootMasterKeepsItsRegisteredPrimitiveIdentity() {
        var document = VisioDocument.Create();
        var blueprint = new VisioShape("root", 1, .5, 2, 1, "") {
            NameU = "Triangle", FillColor = OfficeColor.Blue, LinePattern = 0,
            TextStyle = new VisioTextStyle { Size = 12 }
        };
        var master = document.RegisterMaster("Ellipse", blueprint);
        VisioPage page = document.AddPage("Page", 4, 4);
        VisioShape instance = page.AddShape("instance", master, 2, 2, 2, 1);
        Assert.Same(blueprint, instance.MasterShape);
        Assert.Single(ShapeGroup(page, "instance").Elements(Svg + "ellipse"));
        foreach (VisioDocument saved in new[] { document, Reopen(document, "vsdx"), Reopen(document, "vdx") }) {
            OfficeRasterImage image = Raster(saved.Pages[0]);
            Assert.Equal(OfficeColor.Blue, image.GetPixel(255, 165));
            Assert.Equal(OfficeColor.White, image.GetPixel(105, 155));
        }
    }

    [Fact]
    public void RootStencilSemanticsDoNotReplaceReferencedChildPrimitives() {
        var stencil = new VisioStencilShape("flow.startend", "Start/End", "Ellipse", "Flowchart", 6, 2,
            tags: new[] { "Terminator" });
        var groupDocument = VisioDocument.Create();
        var groupMaster = groupDocument.RegisterMaster("Ellipse", Blueprint());
        groupMaster.StencilTags = new[] { "Terminator" };
        VisioPage groupPage = groupDocument.AddPage("Group", 8, 4);
        groupPage.AddShape("instance", groupMaster, 4, 2, 6, 2);
        AssertSilhouettes(groupPage);

        var rootDocument = VisioDocument.Create();
        VisioPage rootPage = rootDocument.AddPage("Root", 8, 4);
        rootPage.AddStencilShape(stencil, "root", 4, 2);
        XElement root = ShapeGroup(rootPage, "root");
        Assert.Empty(root.Elements(Svg + "ellipse"));
        string outline = Assert.Single(root.Elements(Svg + "path")).Attribute("d")!.Value;
        Assert.StartsWith("M 200 300 L 600 300 L", outline, StringComparison.Ordinal);
        Assert.Contains("L 600 100 L 200 100 L", outline, StringComparison.Ordinal);
    }

    private static VisioShape Blueprint() {
        var root = new VisioShape("root", 3, 1, 6, 2, "") { FillPattern = 0, LinePattern = 0 };
        root.Children.Add(Primitive("ellipse", "Ellipse", 1));
        root.Children.Add(Primitive("diamond", "Diamond", 3));
        var nested = new VisioShape("nested", 5, 1, 2, 2, "") { FillPattern = 0, LinePattern = 0 };
        nested.Children.Add(Primitive("triangle", "Triangle", 1));
        root.Children.Add(nested);
        return root;
    }

    private static VisioShape Primitive(string id, string kind, double x) =>
        new(id, x, 1, 1.5, 1, "") { NameU = kind, FillColor = OfficeColor.Blue, LinePattern = 0 };

    private static void AssertSilhouettes(VisioPage page) {
        XElement ellipse = ShapeGroup(page, "instance:ellipse");
        if (ellipse.Element(Svg + "ellipse") == null) {
            Assert.StartsWith("M 275 200 L", Assert.Single(ellipse.Elements(Svg + "path")).Attribute("d")!.Value,
                StringComparison.Ordinal);
        } else {
            Assert.Single(ellipse.Elements(Svg + "ellipse"));
        }
        Assert.Equal("M 400 250 L 475 200 L 400 150 L 325 200 Z",
            Assert.Single(ShapeGroup(page, "instance:diamond").Elements(Svg + "path")).Attribute("d")!.Value);
        Assert.Equal("M 525 250 L 600 150 L 675 250 Z",
            Assert.Single(ShapeGroup(page, "instance:triangle").Elements(Svg + "path")).Attribute("d")!.Value);
        Assert.Empty(ShapeGroup(page, "instance").Elements(Svg + "path"));
        Assert.Empty(ShapeGroup(page, "instance:nested").Elements(Svg + "path"));

        OfficeRasterImage image = Raster(page);
        Assert.Equal(OfficeColor.Blue, image.GetPixel(242, 165));
        Assert.Equal(OfficeColor.White, image.GetPixel(130, 155));
        Assert.Equal(OfficeColor.Blue, image.GetPixel(400, 200));
        Assert.Equal(OfficeColor.White, image.GetPixel(442, 165));
        Assert.Equal(OfficeColor.Blue, image.GetPixel(600, 200));
        Assert.Equal(OfficeColor.Blue, image.GetPixel(650, 240));
        Assert.Equal(OfficeColor.White, image.GetPixel(650, 160));
    }

    private static XElement ShapeGroup(VisioPage page, string id) => XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions {
        PixelsPerInch = 100, BackgroundColor = null, RenderText = false, RenderStencilArtwork = false
    })).Descendants(Svg + "g").Single(g => (string?)g.Attribute("data-visio-shape-id") == id);

    private static OfficeRasterImage Raster(VisioPage page) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ToPng(new VisioPngSaveOptions {
            PixelsPerInch = 100, BackgroundColor = OfficeColor.White, Supersampling = 1,
            RenderText = false, RenderStencilArtwork = false
        }), out OfficeRasterImage? image));
        return image!;
    }

    private static VisioDocument Reopen(VisioDocument document, string format) => format == "vdx"
        ? VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value
        : VisioDocument.Load(new MemoryStream(document.ToBytes()));
}
