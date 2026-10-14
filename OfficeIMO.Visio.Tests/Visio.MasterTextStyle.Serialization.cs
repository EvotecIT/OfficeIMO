using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioMasterTextStyleSerializationTests {
    private static readonly XNamespace Native = "http://schemas.microsoft.com/office/visio/2012/main";
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";

    [Theory]
    [InlineData(0)] // Character/color styling without frame caches.
    [InlineData(1)] // Explicit local pins with default frame dimensions.
    [InlineData(2)] // Explicit frame dimensions with default local pins.
    [InlineData(3)] // Unstyled simple master keeps its omitted-text behavior.
    public void SimpleMasterFrameDefaultsAgreeBeforeAndAfterReopeningAndResizing(int profile) {
        VisioDocument document = VisioDocument.Create();
        VisioTextStyle? style = profile == 3 ? null : new VisioTextStyle {
            FontFamily = "Arial", Size = 16, Color = OfficeColor.Black, BackgroundColor = OfficeColor.Red,
            HorizontalAlignment = VisioTextHorizontalAlignment.Center, VerticalAlignment = VisioTextVerticalAlignment.Middle,
            LeftMargin = 0, RightMargin = 0, TopMargin = 0, BottomMargin = 0
        };
        if (profile == 1) { style!.TextLocPinX = 0.5; style.TextLocPinY = 0.1; style.TextAngle = Math.PI / 2; }
        if (profile == 2) { style!.TextWidth = 0.8; style.TextHeight = 0.4; }
        var blueprint = new VisioShape("1", 1, 0.5, 2, 1, profile == 3 ? "Blueprint text" : "") { LinePattern = 0, FillPattern = 0, TextStyle = style };
        VisioMaster master = document.RegisterMaster("Default frame", blueprint);
        VisioPage seed = document.AddPage("Seed").Size(8, 8);
        VisioShape blank = seed.AddShape("blank", master, 3, 3, 2, 1);
        Assert.Equal("", blank.Text);
        AssertResolvedFrame(blank.TextStyle, profile, resized: false);
        blank.Text = "MMMMMM";

        XElement emitted = MasterShapes(document.ToBytes())["Default frame"];
        foreach (string name in new[] { "TxtPinX", "TxtPinY", "TxtWidth", "TxtHeight", "TxtLocPinX", "TxtLocPinY", "TxtAngle" })
            Assert.Single(emitted.Elements(Native + "Cell"), cell => (string?)cell.Attribute("N") == name);
        Assert.Equal("Width*0.5", Formula(emitted, "TxtPinX"));
        Assert.Equal("Height*0.5", Formula(emitted, "TxtPinY"));
        Assert.Equal(profile == 2 ? null : "Width*0.875", Formula(emitted, "TxtWidth"));
        Assert.Equal(profile == 2 ? null : "Height*0.75", Formula(emitted, "TxtHeight"));
        Assert.Equal(profile == 1 ? null : "TxtWidth*0.5", Formula(emitted, "TxtLocPinX"));
        Assert.Equal(profile == 1 ? null : "TxtHeight*0.5", Formula(emitted, "TxtLocPinY"));

        int lane = 0;
        foreach (VisioDocument candidate in Reopened(document)) {
            lane++;
            VisioMaster available = candidate.GetMaster("Default frame");
            VisioPage ordinary = candidate.AddPage("Object overload " + lane).Size(8, 8);
            VisioShape normal = ordinary.AddShape("normal", available, 3, 3, 2, 1, text: "MMMMMM");
            VisioPage resized = candidate.AddPage("Name overload " + lane).Size(8, 8);
            VisioShape enlarged = resized.AddShape("enlarged", "Default frame", 5, 5, 4, 3, text: "MMMMMM", unit: VisioMeasurementUnit.Inches);
            AssertResolvedFrame(normal.TextStyle, profile, resized: false);
            AssertResolvedFrame(enlarged.TextStyle, profile, resized: true);
            Assert.NotSame(available.Shape.TextStyle, normal.TextStyle);
            Assert.NotSame(available.Shape.TextStyle, enlarged.TextStyle);
            if (profile == 3) {
                normal.TextStyle!.BackgroundColor = OfficeColor.Red;
                enlarged.TextStyle!.BackgroundColor = OfficeColor.Red;
            }
            AssertRenderedFrame(ordinary, profile == 1 ? 2.725 : 3, profile == 1 ? 3.375 : 3);
            AssertRenderedFrame(resized, profile == 1 ? 4.175 : 5, profile == 1 ? 5.75 : 5);
        }
        Assert.Same(style, blueprint.TextStyle);
        if (style == null) { Assert.Equal("Blueprint text", blueprint.Text); return; }
        Assert.Null(style.TextPinX);
        Assert.Null(style.TextPinY);
        Assert.Equal(profile == 2 ? 0.8 : (double?)null, style.TextWidth);
        Assert.Equal(profile == 2 ? 0.4 : (double?)null, style.TextHeight);
        Assert.Equal(profile == 1 ? 0.5 : (double?)null, style.TextLocPinX);
        Assert.Equal(profile == 1 ? 0.1 : (double?)null, style.TextLocPinY);
        Assert.Equal(profile == 1 ? Math.PI / 2 : (double?)null, style.TextAngle);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StyledRegisteredGeometryAndAliasesRetainTheirMasterIdentityAcrossReopening(bool namedBlueprint) {
        string[] names = VisioDocument.SupportedBuiltinMasters.Where(name => name != "Dynamic connector").ToArray();
        VisioDocument control = VisioDocument.Create();
        VisioDocument styled = VisioDocument.Create();
        VisioPage controlPage = control.AddPage("Control");
        VisioPage page = styled.AddPage("Styled");
        foreach (string name in names) {
            string? sourceName = namedBlueprint ? "Authored blueprint" : null;
            VisioMaster controlMaster = control.RegisterMaster(name, Blueprint(sourceName));
            controlPage.AddShape(name, controlMaster, 2, 2, 2, 1);
            VisioShape blueprint = Blueprint(sourceName);
            blueprint.TextStyle = FrameStyle();
            VisioMaster master = styled.RegisterMaster(name, blueprint);
            VisioShape byMaster = page.AddShape(name + " object", master, 2, 2, 2, 1);
            VisioShape byName = page.AddShape(name + " name", name, 4, 2, 2, 1, text: "Name overload", unit: VisioMeasurementUnit.Inches);
            AssertFrameStyle(byMaster.TextStyle);
            AssertFrameStyle(byName.TextStyle);
            Assert.NotSame(blueprint.TextStyle, byMaster.TextStyle);
            Assert.NotSame(blueprint.TextStyle, byName.TextStyle);
            byMaster.TextStyle!.Size = 18;
            Assert.Equal(16, byName.TextStyle!.Size);
            Assert.Equal(16, blueprint.TextStyle.Size);
            Assert.Equal(sourceName, blueprint.NameU);
        }

        Dictionary<string, XElement> expected = MasterShapes(control.ToBytes());
        foreach (VisioDocument candidate in Reopened(styled)) {
            Dictionary<string, XElement> actual = MasterShapes(candidate.ToBytes());
            foreach (string name in names) {
                XElement master = actual[name];
                Assert.Equal(name, (string?)master.Attribute("NameU"));
                Assert.Equal(Geometry(expected[name]), Geometry(master));
                Assert.Equal(Cell(expected[name], "LockAspect"), Cell(master, "LockAspect"));
                AssertNativeFrameStyle(master);
                AssertFrameStyle(candidate.GetMaster(name).Shape.TextStyle);
                VisioShape nameInstance = candidate.Pages[0].Shapes.Single(shape => shape.Text == "Name overload" && shape.Master?.NameU == name);
                AssertFrameStyle(nameInstance.TextStyle);
            }
        }
        foreach (VisioMaster master in styled.Masters) {
            Assert.Equal(namedBlueprint ? "Authored blueprint" : null, master.Shape.NameU);
            AssertFrameStyle(master.Shape.TextStyle);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StylingAnExistingDynamicMasterRetainsItsOneDimensionalRoutingAndControl(bool frame) {
        VisioDocument document = VisioDocument.Create();
        document.UseMastersByDefault = true;
        VisioPage page = document.AddPage("Connected");
        VisioShape first = page.AddRectangle(1, 2, 2, 1, "First");
        VisioShape second = page.AddRectangle(5, 2, 2, 1, "Second");
        page.AddConnector(first, second, ConnectorKind.Dynamic);
        XElement before = MasterShapes(document.ToBytes())["Dynamic connector"];
        VisioMaster master = document.GetMaster("Dynamic connector");
        master.Shape.TextStyle = frame ? FrameStyle() : new VisioTextStyle { Color = OfficeColor.Red };

        foreach (VisioDocument candidate in Reopened(document)) {
            XElement after = MasterShapes(candidate.ToBytes())["Dynamic connector"];
            foreach (string name in new[] { "OneD", "BeginX", "BeginY", "EndX", "EndY", "ObjType", "LockHeight", "GlueType", "DynFeedback" }) {
                Assert.NotNull(Cell(before, name));
                Assert.Equal(Cell(before, name), Cell(after, name));
            }
            Assert.Equal(Section(before, "Control").ToString(), Section(after, "Control").ToString());
            Assert.Equal("Dynamic connector", (string?)after.Attribute("NameU"));
            VisioConnector connector = Assert.Single(candidate.Pages[0].Connectors);
            Assert.Equal(ConnectorKind.Dynamic, connector.Kind);
            Assert.NotNull(connector.From);
            Assert.NotNull(connector.To);
            if (frame) {
                AssertNativeFrameStyle(after);
                AssertFrameStyle(candidate.GetMaster("Dynamic connector").Shape.TextStyle);
            } else {
                XElement character = Section(after, "Character").Element(Native + "Row")!;
                Assert.Equal("#FF0000", Cell(character, "Color"));
                Assert.Equal(10D / 72, double.Parse(Cell(character, "Size")!, CultureInfo.InvariantCulture), 12);
                Assert.Single(after.Elements(Native + "Section"), section => (string?)section.Attribute("N") == "Character");
            }
        }
        Assert.Equal("Dynamic connector", master.Shape.NameU);
        if (frame) AssertFrameStyle(master.Shape.TextStyle);
        else Assert.Null(master.Shape.TextStyle.Size);
    }

    private static VisioShape Blueprint(string? name) => new("1", 1, 0.5, 2, 1, "Styled master") { NameU = name };

    private static VisioTextStyle FrameStyle() => new() {
        FontFamily = "Arial", Size = 16, Bold = true, Color = OfficeColor.Red,
        HorizontalAlignment = VisioTextHorizontalAlignment.Right, VerticalAlignment = VisioTextVerticalAlignment.Middle,
        LeftMargin = 0.1, RightMargin = 0.2, TopMargin = 0.03, BottomMargin = 0.04,
        TextPinX = 1.5, TextPinY = 0.5, TextWidth = 1, TextHeight = 0.25,
        TextLocPinX = 1, TextLocPinY = 0.125, TextAngle = -Math.PI / 2
    };

    private static void AssertFrameStyle(VisioTextStyle? style) {
        Assert.NotNull(style);
        Assert.Equal("Arial", style.FontFamily);
        Assert.Equal(16, style.Size);
        Assert.True(style.Bold);
        Assert.Equal(OfficeColor.Red, style.Color);
        Assert.Equal(VisioTextHorizontalAlignment.Right, style.HorizontalAlignment);
        Assert.Equal(VisioTextVerticalAlignment.Middle, style.VerticalAlignment);
        Assert.Equal(0.1, style.LeftMargin);
        Assert.Equal(0.2, style.RightMargin);
        Assert.Equal(0.03, style.TopMargin);
        Assert.Equal(0.04, style.BottomMargin);
        Assert.Equal(1.5, style.TextPinX);
        Assert.Equal(0.5, style.TextPinY);
        Assert.Equal(1, style.TextWidth);
        Assert.Equal(0.25, style.TextHeight);
        Assert.Equal(1, style.TextLocPinX);
        Assert.Equal(0.125, style.TextLocPinY);
        Assert.Equal(-Math.PI / 2, style.TextAngle!.Value, 12);
    }

    private static void AssertNativeFrameStyle(XElement master) {
        foreach (var entry in new[] { ("TxtPinX", 1.5), ("TxtPinY", 0.5), ("TxtWidth", 1D), ("TxtHeight", 0.25),
            ("TxtLocPinX", 1D), ("TxtLocPinY", 0.125), ("TxtAngle", -Math.PI / 2), ("LeftMargin", 0.1),
            ("RightMargin", 0.2), ("TopMargin", 0.03), ("BottomMargin", 0.04), ("VerticalAlign", 1D) }) {
            XElement cell = Assert.Single(master.Elements(Native + "Cell"), cell => (string?)cell.Attribute("N") == entry.Item1);
            Assert.Equal(entry.Item2, double.Parse(cell.Attribute("V")!.Value, CultureInfo.InvariantCulture), 12);
        }
        XElement character = Section(master, "Character").Element(Native + "Row")!;
        Assert.Equal(16D / 72, double.Parse(Cell(character, "Size")!, CultureInfo.InvariantCulture), 12);
        Assert.Equal("#FF0000", Cell(character, "Color"));
        Assert.Equal("1", Cell(character, "Style"));
        Assert.NotNull(Cell(character, "Font"));
        Assert.Equal("2", Cell(Section(master, "Paragraph").Element(Native + "Row")!, "HorzAlign"));
    }

    private static XElement Section(XElement shape, string name) => Assert.Single(shape.Elements(Native + "Section"), section => (string?)section.Attribute("N") == name);

    private static void AssertResolvedFrame(VisioTextStyle? style, int profile, bool resized) {
        Assert.NotNull(style);
        if (profile != 3) Assert.Equal(16, style.Size);
        Assert.Equal(profile == 2 ? resized ? 1.6 : 0.8 : resized ? 3.5 : 1.75, style.TextWidth);
        Assert.Equal(profile == 2 ? resized ? 1.2 : 0.4 : resized ? 2.25 : 0.75, style.TextHeight!.Value, 12);
        Assert.Equal(resized ? 2 : 1, style.TextPinX);
        Assert.Equal(resized ? 1.5 : 0.5, style.TextPinY);
        Assert.Equal(profile == 1 ? resized ? 1 : 0.5 : profile == 2 ? resized ? 0.8 : 0.4 : resized ? 1.75 : 0.875, style.TextLocPinX);
        Assert.Equal(profile == 1 ? resized ? 0.3 : 0.1 : profile == 2 ? resized ? 0.6 : 0.2 : resized ? 1.125 : 0.375, style.TextLocPinY!.Value, 12);
        Assert.Equal(profile == 1 ? Math.PI / 2 : 0, style.TextAngle!.Value, 12);
    }

    private static void AssertRenderedFrame(VisioPage page, double centerX, double centerY) {
        byte[] before = page.OwnerDocument!.ToLegacyXmlResult().Value;
        const double scale = 100;
        XDocument svg = XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = scale, RenderStencilArtwork = false }));
        XElement background = Assert.Single(svg.Descendants(Svg + "rect"), rectangle => rectangle.Attribute("data-officeimo-text-background") != null);
        double Read(string name) => double.Parse(background.Attribute(name)!.Value, CultureInfo.InvariantCulture);
        Assert.Equal(centerX * scale, Read("x") + Read("width") / 2, 3);
        Assert.Equal((page.Height - centerY) * scale, Read("y") + Read("height") / 2, 3);
        OfficeRasterImage raster = VisualBaselineTestSupport.DecodePng(page.ToPng(new VisioPngSaveOptions {
            PixelsPerInch = scale, Supersampling = 1, RenderStencilArtwork = false
        }), "Default-frame PNG could not be decoded.");
        int left = raster.Width, top = raster.Height, right = -1, bottom = -1;
        for (int y = 0; y < raster.Height; y++) {
            for (int x = 0; x < raster.Width; x++) {
                OfficeColor color = raster.GetPixel(x, y);
                if (color.R > 220 && color.G < 40 && color.B < 40) {
                    left = Math.Min(left, x); right = Math.Max(right, x);
                    top = Math.Min(top, y); bottom = Math.Max(bottom, y);
                }
            }
        }
        Assert.True(right >= left && bottom >= top, "Expected a visible text frame background.");
        Assert.InRange((left + right + 1) / 2D, centerX * scale - 1, centerX * scale + 1);
        Assert.InRange((top + bottom + 1) / 2D, (page.Height - centerY) * scale - 1, (page.Height - centerY) * scale + 1);
        Assert.Equal(before, page.OwnerDocument.ToLegacyXmlResult().Value);
    }

    private static string? Formula(XElement shape, string name) => shape.Elements(Native + "Cell").Single(cell => (string?)cell.Attribute("N") == name).Attribute("F")?.Value;

    private static string? Cell(XElement shape, string name) => shape.Elements(Native + "Cell").SingleOrDefault(cell => (string?)cell.Attribute("N") == name)?.Attribute("V")?.Value;

    private static string[] Geometry(XElement shape) => Section(shape, "Geometry").Elements(Native + "Row")
        .Where(row => (string?)row.Attribute("T") != "Geometry")
        .Select(row => (string?)row.Attribute("T") + ":" + string.Join(";", row.Elements(Native + "Cell").Select(cell => (string?)cell.Attribute("N") + "=" + (string?)cell.Attribute("V")))).ToArray();

    private static Dictionary<string, XElement> MasterShapes(byte[] bytes) {
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        XNamespace relation = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XDocument catalog = Read("visio/masters/masters.xml");
        XDocument relationships = Read("visio/masters/_rels/masters.xml.rels");
        var shapes = new Dictionary<string, XElement>();
        foreach (XElement master in catalog.Root!.Elements(Native + "Master")) {
            string id = master.Element(Native + "Rel")!.Attribute(relation + "id")!.Value;
            string target = relationships.Root!.Elements().Single(entry => (string?)entry.Attribute("Id") == id).Attribute("Target")!.Value;
            shapes.Add(master.Attribute("NameU")!.Value, Read("visio/masters/" + target).Root!.Element(Native + "Shapes")!.Element(Native + "Shape")!);
        }
        return shapes;

        XDocument Read(string path) {
            ZipArchiveEntry? entry = zip.GetEntry(path);
            Assert.NotNull(entry);
            using Stream stream = entry.Open();
            return XDocument.Load(stream);
        }
    }

    private static IEnumerable<VisioDocument> Reopened(VisioDocument document) {
        yield return document;
        yield return VisioDocument.Load(new MemoryStream(document.ToBytes()));
        yield return VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value;
    }
}
