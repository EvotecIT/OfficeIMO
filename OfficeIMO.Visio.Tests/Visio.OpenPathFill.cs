using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioOpenPathFillTests {
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";

    [Theory]
    [InlineData(false, "direct")]
    [InlineData(false, "vsdx")]
    [InlineData(false, "vdx")]
    [InlineData(true, "direct")]
    [InlineData(true, "vsdx")]
    [InlineData(true, "vdx")]
    public void OpenNurbsKeepsItsFillWithoutStrokingTheClosingChord(bool resized, string route) {
        var source = CreateDocument(Section(false, Point("MoveTo", 0, 0), Nurbs()));
        if (resized) source.Pages[0].ResizeShape(source.Pages[0].Shapes[0], 2, 2);
        var document = Reopen(source, route);
        string geometry = SavedGeometry(document);

        AssertOpenCurve(document.Pages[0], resized);

        Assert.Equal(geometry, SavedGeometry(document));
    }

    [Theory]
    [InlineData("straight")]
    [InlineData("cubic")]
    public void FillClosureDoesNotDependOnTheOpenSegmentFamily(string family) {
        XElement[] rows = family == "straight"
            ? new[] { Point("MoveTo", 0, 0), Point("LineTo", 1, 0), Point("LineTo", 1, 1) }
            : new[] { Point("MoveTo", 0, 0), Point("CubBezTo", 1, 1,
                ("A", 1D / 3), ("B", 8D / 15), ("C", 2D / 3), ("D", 13D / 15)) };
        var document = CreateDocument(Section(false, rows));
        XElement path = Assert.Single(Paths(document.Pages[0]));
        Assert.False(IsClosed(path));
        Assert.Equal("#FF0000", (string?)path.Attribute("fill"));
        OfficeRasterImage image = Raster(document.Pages[0]);
        AssertColor(image, 149, family == "straight" ? 251 : 248, OfficeColor.Red);
        AssertColor(image, family == "straight" ? 199 : 150, family == "straight" ? 250 : 235, OfficeColor.Blue);
    }

    [Fact]
    public void MixedContoursShareTheirFillGroupWithoutSharingStrokeClosure() {
        XElement strokeOnly = Point("MoveTo", .2, .2);
        strokeOnly.Add(Cell("NoFill", "1"));
        XElement inner = Point("MoveTo", .35, .35);
        inner.Add(Cell("NoLine", "1"));
        var document = CreateDocument(Section(false,
            Point("MoveTo", .1, .1), Point("LineTo", .9, .1), Point("LineTo", .9, .9),
            Point("LineTo", .1, .9), Point("LineTo", .1, .1),
            strokeOnly, Point("LineTo", .8, .2),
            inner, Point("LineTo", .65, .35), Point("LineTo", .65, .65)));
        List<XElement> paths = Paths(document.Pages[0]);
        XElement fill = Assert.Single(paths, p => (string?)p.Attribute("fill") != "none");
        Assert.Equal("evenodd", (string?)fill.Attribute("fill-rule"));
        Assert.Equal("none", (string?)fill.Attribute("stroke"));
        XElement[] strokes = paths.Where(p => (string?)p.Attribute("stroke") != "none").ToArray();
        Assert.Equal(2, strokes.Length);
        Assert.Single(strokes, IsClosed);
        Assert.Single(strokes, p => !IsClosed(p));
        OfficeRasterImage image = Raster(document.Pages[0]);
        AssertColor(image, 120, 250, OfficeColor.Red);
        AssertColor(image, 155, 255, OfficeColor.White);
        AssertColor(image, 164, 250, OfficeColor.White);
        AssertColor(image, 110, 250, OfficeColor.Blue);
        AssertColor(image, 150, 280, OfficeColor.Blue);
    }

    [Theory]
    [InlineData("explicit", true)]
    [InlineData("ellipse", true)]
    [InlineData("infinite", false)]
    [InlineData("two-point", false)]
    public void GenuineClosureAndDegenerateFillGuardsRemainIndependent(string kind, bool closed) {
        XElement[] rows = kind switch {
            "ellipse" => new[] { Point("Ellipse", .5, .5, ("A", 1), ("B", .5), ("C", .5), ("D", 1)) },
            "infinite" => new[] { Point("InfiniteLine", 0, .5, ("A", 1), ("B", .5)) },
            "two-point" => new[] { Point("MoveTo", 0, 0), Point("LineTo", 1, 1) },
            _ => new[] { Point("MoveTo", 0, 0), Point("LineTo", 1, 0), Point("LineTo", 1, 1), Point("LineTo", 0, 0) }
        };
        var document = CreateDocument(Section(kind != "two-point", rows));
        XElement path = Assert.Single(Paths(document.Pages[0]));
        Assert.Equal(closed, IsClosed(path));
        Assert.Equal("none", (string?)path.Attribute("fill"));
        OfficeRasterImage image = Raster(document.Pages[0]);
        if (kind == "ellipse") AssertColor(image, 199, 255, OfficeColor.Blue);
        else AssertColor(image, 149, kind == "infinite" ? 250 : 249, OfficeColor.Blue);
        AssertColor(image, 150, 260, OfficeColor.White);
    }

    [Fact]
    public void InheritedOpenNurbsKeepsItsClosureDuringScalingAndReopening() {
        const string xml = "<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'>" +
            "<Masters><Master ID='1' NameU='Curve'><Shapes><Shape ID='1'><XForm><Width>1</Width><Height>1</Height>" +
            "<LocPinX>0.5</LocPinX><LocPinY>0.5</LocPinY></XForm><Geom IX='0'><NoFill>0</NoFill><NoLine>0</NoLine>" +
            "<MoveTo IX='1'><X>0</X><Y>0</Y></MoveTo><NURBSTo IX='2'><X>1</X><Y>1</Y><A>1</A><B>1</B><C>0</C><D>1</D>" +
            "<E>NURBS(2,2,1,1,0.5,0.8,1,1)</E></NURBSTo></Geom></Shape></Shapes></Master></Masters>" +
            "<Pages><Page ID='0' Name='Page'><PageSheet><PageProps><PageWidth>4</PageWidth><PageHeight>4</PageHeight></PageProps>" +
            "</PageSheet><Shapes><Shape ID='2' Master='1'><XForm><PinX>1.5</PinX><PinY>1.5</PinY><Width>2</Width><Height>2</Height>" +
            "<LocPinX>1</LocPinX><LocPinY>1</LocPinY></XForm></Shape></Shapes></Page></Pages></VisioDocument>";
        var source = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml))).Value;
        Style(source.Pages[0].Shapes[0]);
        Assert.Empty(source.Pages[0].Shapes[0].PreservedGeometrySections);
        AssertOpenCurve(source.Pages[0], resized: true);
        foreach (string route in new[] { "vsdx", "vdx" }) {
            var reopened = Reopen(source, route);
            AssertOpenCurve(reopened.Pages[0], resized: true);
        }
    }

    [Fact]
    public void RenderingLeavesNativeConnectorStructuralPointsAndSavedGeometryUnchanged() {
        const string xml = "<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'><Pages><Page ID='0' Name='Page'>" +
            "<PageSheet><PageProps><PageWidth>4</PageWidth><PageHeight>4</PageHeight></PageProps></PageSheet><Shapes><Shape ID='1'>" +
            "<XForm><PinX>1.5</PinX><PinY>1.5</PinY><Width>1</Width><Height>1</Height><LocPinX>0.5</LocPinX><LocPinY>0.5</LocPinY></XForm>" +
            "<XForm1D><BeginX>1</BeginX><BeginY>1</BeginY><EndX>2</EndX><EndY>2</EndY></XForm1D><Geom IX='0'><NoFill>0</NoFill>" +
            "<MoveTo IX='1'><X>0</X><Y>0</Y></MoveTo><NURBSTo IX='2'><X>1</X><Y>1</Y><A>1</A><B>1</B><C>0</C><D>1</D>" +
            "<E>NURBS(2,2,1,1,0.5,0.8,1,1)</E></NURBSTo></Geom></Shape></Shapes></Page></Pages></VisioDocument>";
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml))).Value;
        VisioConnector connector = Assert.Single(document.Pages[0].Connectors);
        var before = VisioConnectorGeometry.GetPoints(connector).ToArray();
        string geometry = SavedGeometry(document);
        Assert.Equal(new OfficePoint(1, 1), connector.StartPoint);
        Assert.Equal(new OfficePoint(2, 2), connector.EndPoint);
        document.Pages[0].ToSvg(new VisioSvgSaveOptions { RenderText = false, RenderStencilArtwork = false });
        Raster(document.Pages[0]);
        Assert.Equal(before, VisioConnectorGeometry.GetPoints(connector));
        Assert.Equal(geometry, SavedGeometry(document));
    }

    private static void AssertOpenCurve(VisioPage page, bool resized) {
        XElement path = Assert.Single(Paths(page));
        Assert.False(IsClosed(path));
        Assert.Equal("#FF0000", (string?)path.Attribute("fill"));
        Assert.Equal("#0000FF", (string?)path.Attribute("stroke"));
        OfficeRasterImage image = Raster(page);
        AssertColor(image, 149, 248, OfficeColor.Red);
        AssertColor(image, 150, 242, OfficeColor.Red);
        AssertColor(image, 150, 260, OfficeColor.White);
        AssertColor(image, 150, resized ? 220 : 235, OfficeColor.Blue);
    }

    private static List<XElement> Paths(VisioPage page) => XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions {
        PixelsPerInch = 100, BackgroundColor = null, RenderText = false, RenderStencilArtwork = false
    })).Descendants().Where(e => (string?)e.Attribute("data-officeimo-preserved-geometry") == "true").ToList();

    private static bool IsClosed(XElement path) => path.Attribute("d")!.Value.EndsWith(" Z", StringComparison.Ordinal);

    private static OfficeRasterImage Raster(VisioPage page) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(page.ToPng(new VisioPngSaveOptions {
            PixelsPerInch = 100, BackgroundColor = OfficeColor.White, Supersampling = 1,
            RenderText = false, RenderStencilArtwork = false
        }), out OfficeRasterImage? image));
        return image!;
    }

    private static void AssertColor(OfficeRasterImage image, int x, int y, OfficeColor expected) {
        OfficeColor actual = image.GetPixel(x, y);
        Assert.InRange(Math.Abs(actual.R - expected.R), 0, 8);
        Assert.InRange(Math.Abs(actual.G - expected.G), 0, 8);
        Assert.InRange(Math.Abs(actual.B - expected.B), 0, 8);
        Assert.Equal(255, actual.A);
    }

    private static VisioDocument CreateDocument(XElement geometry) {
        var source = VisioDocument.Create();
        VisioShape shape = source.AddPage("Page", 4, 4).AddRectangle(1.5, 1.5, 1, 1);
        Style(shape);
        shape.PreservedGeometrySections.Clear();
        shape.PreservedGeometrySections.Add(geometry);
        return VisioDocument.Load(new MemoryStream(source.ToBytes()));
    }

    private static void Style(VisioShape shape) {
        shape.FillColor = OfficeColor.Red; shape.FillPattern = 1;
        shape.LineColor = OfficeColor.Blue; shape.LinePattern = 1; shape.LineWeight = .04;
    }

    private static VisioDocument Reopen(VisioDocument document, string route) => route == "direct" ? document : route == "vdx"
        ? VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value
        : VisioDocument.Load(new MemoryStream(document.ToBytes()));

    private static string SavedGeometry(VisioDocument document) {
        using var zip = new ZipArchive(new MemoryStream(document.ToBytes()));
        using var stream = zip.GetEntry("visio/pages/page1.xml")!.Open();
        return string.Join("\n", XDocument.Load(stream).Descendants(Modern + "Section")
            .Where(s => (string?)s.Attribute("N") == "Geometry").Select(s => s.ToString(SaveOptions.DisableFormatting)));
    }

    private static XElement Section(bool noFill, params XElement[] rows) {
        for (int i = 0; i < rows.Length; i++) rows[i].SetAttributeValue("IX", i + 1);
        return new XElement(Modern + "Section", new XAttribute("N", "Geometry"), new XAttribute("IX", 0),
            Cell("NoFill", noFill ? "1" : "0"), Cell("NoLine", "0"), rows);
    }

    private static XElement Point(string type, double x, double y, params (string Name, double Value)[] cells) =>
        new(Modern + "Row", new XAttribute("T", type), Cell("X", x.ToString("R", CultureInfo.InvariantCulture)),
            Cell("Y", y.ToString("R", CultureInfo.InvariantCulture)), cells.Select(c => Cell(c.Name, c.Value.ToString("R", CultureInfo.InvariantCulture))));

    private static XElement Nurbs() {
        XElement row = Point("NURBSTo", 1, 1, ("A", 1), ("B", 1), ("C", 0), ("D", 1));
        row.Add(Cell("E", "NURBS(2,2,1,1,0.5,0.8,1,1)"));
        return row;
    }

    private static XElement Cell(string name, string value) => new(Modern + "Cell", new XAttribute("N", name), new XAttribute("V", value));
}
