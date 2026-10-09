using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioGeometrySectionContractTests {
    private static readonly XNamespace Native = "http://schemas.microsoft.com/office/visio/2012/main";
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GeneratedShapeAndMasterGeometryHasNativeSectionCellsAndIndexedPathRows(bool masters) {
        string[] kinds = { "Rectangle", "Ellipse", "Diamond", "Triangle", "Pentagon", "Parallelogram", "Hexagon", "Trapezoid", "Off-page reference" };
        VisioDocument document = VisioDocument.Create();
        document.UseMastersByDefault = masters;
        VisioPage page = document.AddPage("Generated").Size(4, 3);
        foreach (string kind in kinds) {
            if (masters) document.RegisterMaster(kind, new VisioShape("blueprint", 1, 1, 2, 1, string.Empty));
            page.Shapes.Add(new VisioShape(kind, 2, 1.5, 2, 1, string.Empty) { NameU = kind });
        }

        XElement[] sections = GetGeometry(document.ToBytes(), masters ? "visio/masters/" : "visio/pages/");
        Assert.Equal(kinds.Length, sections.Length);
        foreach (XElement section in sections) AssertGeneratedSection(section, noFill: false);

        // Legacy export and reopening must preserve the same path identities rather than repair
        // malformed generated Open XML rows only in the conversion bridge.
        VisioDocument xml = VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value;
        XElement[] reopened = GetGeometry(xml.ToBytes(), masters ? "visio/masters/" : "visio/pages/");
        Assert.Equal(sections.Length, reopened.Length);
        for (int i = 0; i < sections.Length; i++) Assert.Equal(RowIdentities(sections[i]), RowIdentities(reopened[i]));
    }

    [Fact]
    public void GeneratedConnectorGeometryUsesSectionFlagsAndStableIndexedRouteRows() {
        VisioDocument document = VisioDocument.Create();
        document.UseMastersByDefault = false;
        VisioPage page = document.AddPage("Routes").Size(4, 3);
        foreach (ConnectorKind kind in new[] { ConnectorKind.Straight, ConnectorKind.RightAngle, ConnectorKind.Dynamic }) {
            VisioConnector connector = page.AddConnector(kind.ToString(), new OfficePoint(.5, .5), new OfficePoint(3.5, 2.5), kind);
            if (kind == ConnectorKind.RightAngle) connector.Waypoints.Add(new VisioConnectorWaypoint(.5, 2.5));
        }
        XElement[] sections = GetGeometry(document.ToBytes(), "visio/pages/");
        Assert.Equal(3, sections.Length);
        foreach (XElement section in sections) AssertGeneratedSection(section, noFill: true);
        VisioDocument reopened = VisioDocument.Load(new MemoryStream(document.ToBytes()));
        XElement[] savedAgain = GetGeometry(reopened.ToBytes(), "visio/pages/");
        for (int i = 0; i < sections.Length; i++) Assert.True(XNode.DeepEquals(sections[i], savedAgain[i]));
    }

    [Theory]
    [InlineData("control")]
    [InlineData("NoShow")]
    [InlineData("NoFill")]
    [InlineData("NoLine")]
    public void StandardGeometryFlagsSurviveCopyAndReopeningAndControlSvgAndPng(string flag) {
        VisioDocument source = LoadXml(ShapeXml(flag));
        var routes = new[] { source, LoadXml(source.ToLegacyXmlResult().Value), VisioDocument.Load(new MemoryStream(source.ToBytes())) };
        foreach (VisioDocument document in routes) {
            AssertShapeRendering(document, document.Pages[0], flag);
            byte[] before = document.ToLegacyXmlResult().Value;
            document.DuplicatePage(document.Pages[0], "Copy");
            AssertShapeRendering(document, document.Pages[1], flag);
            XDocument saved = XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
            XElement[] geometries = saved.Descendants(Legacy + "Geom").ToArray();
            Assert.Equal(2, geometries.Length);
            Assert.True(XNode.DeepEquals(geometries[0], geometries[1]));
            Assert.Equal(new[] { "5", "9", "12", "17", "21" }, geometries[0].Elements().Where(e => e.Attribute("IX") != null).Select(e => (string?)e.Attribute("IX")));
            Assert.Equal("Width*0", (string?)geometries[0].Element(Legacy + "MoveTo")!.Element(Legacy + "X")!.Attribute("F"));
            Assert.Equal("IN", (string?)geometries[0].Element(Legacy + "MoveTo")!.Element(Legacy + "X")!.Attribute("Unit"));
            Assert.True(XNode.DeepEquals(XDocument.Load(new MemoryStream(before)).Descendants(Legacy + "Geom").Single(), geometries[0]));
        }
    }

    [Fact]
    public void HiddenGeometryInheritedFromLoadedMasterStaysHiddenWhenResizedAndCopied() {
        XDocument xml = XDocument.Parse(ShapeXml("NoShow"));
        XElement shape = xml.Descendants(Legacy + "Shape").Single();
        shape.Remove();
        xml.Root!.AddFirst(new XElement(Legacy + "Masters", new XElement(Legacy + "Master", new XAttribute("ID", "1"),
            new XAttribute("NameU", "Hidden"), new XElement(Legacy + "Shapes", shape))));
        xml.Descendants(Legacy + "Page").Single().Element(Legacy + "Shapes")!.Add(XElement.Parse(
            "<Shape xmlns='" + Legacy + "' ID='2' Master='1' Type='Shape'><XForm><PinX>2</PinX><PinY>1.5</PinY><Width>3</Width><Height>2</Height><LocPinX>1.5</LocPinX><LocPinY>1</LocPinY></XForm></Shape>"));
        VisioDocument source = LoadXml(xml.ToString());
        foreach (VisioDocument document in new[] { source, LoadXml(source.ToLegacyXmlResult().Value), VisioDocument.Load(new MemoryStream(source.ToBytes())) }) {
            AssertShapeRendering(document, document.Pages[0], "NoShow");
            document.DuplicatePage(document.Pages[0], "Copy");
            AssertShapeRendering(document, document.Pages[1], "NoShow");
        }
    }

    [Theory]
    [InlineData("NoShow", false)]
    [InlineData("NoLine", false)]
    [InlineData("NoShow", true)]
    [InlineData("NoLine", true)]
    public void HiddenNativeConnectorRetainsCurvedRouteAndLabelWithoutDrawingFallbackLine(string flag, bool historicalHeader) {
        VisioDocument control = LoadXml(ConnectorXml("control"));
        VisioConnector controlConnector = Assert.Single(control.Pages[0].Connectors);
        controlConnector.PlaceLabel(.5, width: .4, height: .2);
        var expected = Assert.Single(control.CreateInspectionSnapshot().Pages[0].Connectors);
        Assert.True(Math.Abs(expected.LabelResolvedPinY!.Value - 1) > .1);

        VisioDocument source = LoadXml(ConnectorXml(flag));
        if (historicalHeader) source = WithHistoricalHeader(source);
        foreach (VisioDocument document in new[] { source, LoadXml(source.ToLegacyXmlResult().Value), VisioDocument.Load(new MemoryStream(source.ToBytes())) }) {
            VisioConnector connector = Assert.Single(document.Pages[0].Connectors);
            connector.PlaceLabel(.5, width: .4, height: .2);
            connector.EndArrow = EndArrow.Triangle;
            var snapshot = Assert.Single(document.CreateInspectionSnapshot().Pages[0].Connectors);
            Assert.Equal(expected.LabelResolvedPinX!.Value, snapshot.LabelResolvedPinX!.Value, 8);
            Assert.Equal(expected.LabelResolvedPinY!.Value, snapshot.LabelResolvedPinY!.Value, 8);
            byte[] before = document.ToLegacyXmlResult().Value;
            AssertHiddenConnectorRendering(document.Pages[0]);
            Assert.Equal(before, document.ToLegacyXmlResult().Value);
            document.DuplicatePage(document.Pages[0], "Copy");
            AssertHiddenConnectorRendering(document.Pages[1]);
            var copied = Assert.Single(document.CreateInspectionSnapshot().Pages[1].Connectors);
            Assert.Equal(snapshot.LabelResolvedPinY!.Value, copied.LabelResolvedPinY!.Value, 8);
        }
    }

    [Fact]
    public void HiddenConnectorOutlineDoesNotDisplaceAnotherConnectorsLabel() {
        XDocument xml = XDocument.Parse(ConnectorXml("NoShow"));
        XElement arc = xml.Descendants(Legacy + "ArcTo").Single();
        arc.Name = Legacy + "LineTo";
        arc.Element(Legacy + "A")!.Remove();
        VisioDocument document = LoadXml(xml.ToString());
        VisioPage page = document.Pages[0];
        page.Connectors[0].Label = null;
        VisioConnector labelled = page.AddConnector("label", new OfficePoint(2, .25), new OfficePoint(2, 1.75));
        labelled.LinePattern = 0;
        labelled.Label = "X";
        labelled.PlaceLabel(.5, width: .6, height: .2);
        XDocument svg = XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 48, ResolveConnectorLabelOverlaps = true }));
        XElement background = Assert.Single(svg.Descendants(Svg + "rect"), e => (string?)e.Attribute("data-officeimo-connector-label-background") == "true");
        // SVG coordinates are rounded to three decimals, including the two independent box edges.
        Assert.InRange(double.Parse(background.Attribute("x")!.Value, CultureInfo.InvariantCulture) + double.Parse(background.Attribute("width")!.Value, CultureInfo.InvariantCulture) / 2, 95.999, 96.001);
        Assert.InRange(double.Parse(background.Attribute("y")!.Value, CultureInfo.InvariantCulture) + double.Parse(background.Attribute("height")!.Value, CultureInfo.InvariantCulture) / 2, 95.999, 96.001);
    }

    [Fact]
    public void NoLineCannotRestoreFallbackStrokeWhenIndexedNativePathRowsAreDeleted() {
        XDocument xml = XDocument.Parse(ConnectorXml("NoLine"));
        foreach (XElement row in xml.Descendants(Legacy + "Geom").Single().Elements().Where(e => e.Attribute("IX") != null))
            row.SetAttributeValue("Del", "1");
        VisioDocument source = LoadXml(xml.ToString());
        foreach (VisioDocument document in new[] { source, VisioDocument.Load(new MemoryStream(source.ToBytes())) }) {
            byte[] before = document.ToLegacyXmlResult().Value;
            AssertHiddenConnectorRendering(document.Pages[0]);
            Assert.Equal(before, document.ToLegacyXmlResult().Value);
            XElement geometry = XDocument.Load(new MemoryStream(before)).Descendants(Legacy + "Geom").Single();
            Assert.Equal(new[] { "5", "9" }, geometry.Elements().Where(e => e.Attribute("IX") != null).Select(e => (string?)e.Attribute("IX")));
            Assert.All(geometry.Elements().Where(e => e.Attribute("IX") != null), row => Assert.Equal("1", (string?)row.Attribute("Del")));
        }
    }

    [Theory]
    [InlineData("NoShow")]
    [InlineData("NoLine")]
    public void HiddenAuxiliaryGeometryDoesNotReplaceTheSingleVisibleNativeConnectorRoute(string flag) {
        VisioDocument control = LoadXml(ConnectorXml("control"));
        control.Pages[0].Connectors[0].PlaceLabel(.5, width: .4, height: .2);
        control.Pages[0].Connectors[0].EndArrow = EndArrow.Triangle;
        double expectedY = control.CreateInspectionSnapshot().Pages[0].Connectors[0].LabelResolvedPinY!.Value;
        VisioSvgSaveOptions svgOptions = new() { PixelsPerInch = 48, RenderStencilArtwork = false, ResolveConnectorLabelOverlaps = false };
        VisioPngSaveOptions pngOptions = new() { PixelsPerInch = 48, Supersampling = 1, RenderStencilArtwork = false, ResolveConnectorLabelOverlaps = false };
        string[] expectedPaths = SvgPaths(control.Pages[0]);
        byte[] expectedPng = control.Pages[0].ToPng(pngOptions);
        XDocument xml = XDocument.Parse(ConnectorXml("control"));
        XElement auxiliary = XElement.Parse("<Geom xmlns='" + Legacy + "' IX='1'><NoFill>1</NoFill><NoLine>" + (flag == "NoLine" ? "1" : "0") + "</NoLine><NoShow>" + (flag == "NoShow" ? "1" : "0") + "</NoShow>" +
            "<MoveTo IX='2'><X>0</X><Y>1</Y></MoveTo><LineTo IX='4'><X>3</X><Y>1</Y></LineTo></Geom>");
        xml.Descendants(Legacy + "Shape").Single().Add(auxiliary);
        VisioDocument source = LoadXml(xml.ToString());
        foreach (VisioDocument document in new[] { source, WithHistoricalHeader(source), LoadXml(source.ToLegacyXmlResult().Value), VisioDocument.Load(new MemoryStream(source.ToBytes())) }) {
            document.Pages[0].Connectors[0].PlaceLabel(.5, width: .4, height: .2);
            document.Pages[0].Connectors[0].EndArrow = EndArrow.Triangle;
            document.DuplicatePage(document.Pages[0], "Copy");
            byte[] before = document.ToLegacyXmlResult().Value;
            foreach (var page in document.CreateInspectionSnapshot().Pages)
                Assert.Equal(expectedY, Assert.Single(page.Connectors).LabelResolvedPinY!.Value, 8);
            foreach (VisioPage page in document.Pages) {
                Assert.Equal(expectedPaths, SvgPaths(page));
                Assert.Equal(expectedPng, page.ToPng(pngOptions));
            }
            Assert.Equal(before, document.ToLegacyXmlResult().Value);
        }

        string[] SvgPaths(VisioPage page) => XDocument.Parse(page.ToSvg(svgOptions)).Descendants(Svg + "path")
            .Select(path => path.Attribute("d")!.Value).ToArray();
    }

    private static void AssertGeneratedSection(XElement section, bool noFill) {
        Assert.Equal("0", (string?)section.Attribute("IX"));
        foreach (string flag in new[] { "NoFill", "NoLine", "NoShow", "NoSnap", "NoQuickDrag" }) {
            XElement cell = Assert.Single(section.Elements(Native + "Cell"), e => (string?)e.Attribute("N") == flag);
            Assert.Equal(flag == "NoFill" && noFill ? "1" : "0", (string?)cell.Attribute("V"));
        }
        XElement[] rows = section.Elements(Native + "Row").ToArray();
        Assert.Equal("MoveTo", (string?)rows[0].Attribute("T"));
        Assert.All(rows.Skip(1), row => Assert.Equal("LineTo", (string?)row.Attribute("T")));
        Assert.Equal(Enumerable.Range(1, rows.Length).Select(i => i.ToString(CultureInfo.InvariantCulture)), rows.Select(row => (string?)row.Attribute("IX")));
    }

    private static string[] RowIdentities(XElement section) => section.Elements(Native + "Row")
        .Select(row => (string?)row.Attribute("IX") + ":" + (string?)row.Attribute("T")).ToArray();

    private static XElement[] GetGeometry(byte[] bytes, string prefix) {
        using ZipArchive zip = new(new MemoryStream(bytes), ZipArchiveMode.Read);
        return zip.Entries.Where(e => e.FullName.StartsWith(prefix, StringComparison.Ordinal) && e.FullName.EndsWith(".xml", StringComparison.Ordinal))
            .SelectMany(e => { using Stream stream = e.Open(); return XDocument.Load(stream).Descendants(Native + "Section").Where(s => (string?)s.Attribute("N") == "Geometry").ToArray(); }).ToArray();
    }

    private static void AssertShapeRendering(VisioDocument document, VisioPage page, string flag) {
        byte[] before = document.ToLegacyXmlResult().Value;
        XDocument svg = XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 48, RenderText = false, RenderStencilArtwork = false }));
        XElement[] paths = svg.Descendants(Svg + "path").ToArray();
        if (flag == "NoShow") Assert.Empty(paths);
        else {
            Assert.NotEmpty(paths);
            Assert.Contains(paths, p => (string?)p.Attribute("fill") == (flag == "NoFill" ? "none" : "#FF0000"));
            Assert.All(paths, p => Assert.Equal(flag == "NoLine" ? "none" : "#0000FF", (string?)p.Attribute("stroke")));
        }
        (int red, int blue) = CountPixels(page, labels: false);
        Assert.Equal(flag != "NoShow" && flag != "NoFill", red > 0);
        Assert.Equal(flag != "NoShow" && flag != "NoLine", blue > 0);
        Assert.Equal(before, document.ToLegacyXmlResult().Value);
    }

    private static void AssertHiddenConnectorRendering(VisioPage page) {
        XDocument svg = XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 48, RenderStencilArtwork = false, ResolveConnectorLabelOverlaps = false }));
        Assert.All(svg.Descendants(Svg + "path"), path => Assert.Equal("none", (string?)path.Attribute("stroke")));
        Assert.Contains(svg.Descendants(Svg + "text"), text => text.Value == "X");
        Assert.Equal(0, CountPixels(page, labels: true).Blue);
    }

    private static (int Red, int Blue) CountPixels(VisioPage page, bool labels) {
        byte[] png = page.ToPng(new VisioPngSaveOptions { PixelsPerInch = 48, Supersampling = 1, RenderText = labels, RenderConnectorLabels = labels, RenderStencilArtwork = false, ResolveConnectorLabelOverlaps = false });
        Assert.True(OfficeRasterImageDecoder.TryDecode(png, out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.Equal(192, image!.Width);
        Assert.Equal(144, image.Height);
        int red = 0, blue = 0;
        for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++) {
            OfficeColor color = image.GetPixel(x, y);
            if (color.R > 200 && color.G < 30 && color.B < 30) red++;
            if (color.B > 200 && color.R < 30 && color.G < 30) blue++;
        }
        return (red, blue);
    }

    private static VisioDocument WithHistoricalHeader(VisioDocument document) {
        using MemoryStream stream = new();
        byte[] bytes = document.ToBytes(); stream.Write(bytes, 0, bytes.Length); stream.Position = 0;
        using (ZipArchive zip = new(stream, ZipArchiveMode.Update, true)) {
            ZipArchiveEntry entry = zip.GetEntry("visio/pages/page1.xml")!;
            XDocument part; using (Stream reader = entry.Open()) part = XDocument.Load(reader);
            foreach (XElement section in part.Descendants(Native + "Section").Where(s => (string?)s.Attribute("N") == "Geometry")) {
                XElement[] cells = section.Elements(Native + "Cell").ToArray(); cells.Remove();
                section.AddFirst(new XElement(Native + "Row", new XAttribute("T", "Geometry"), cells));
            }
            entry.Delete(); using Stream writer = zip.CreateEntry("visio/pages/page1.xml").Open(); part.Save(writer);
        }
        return VisioDocument.Load(new MemoryStream(stream.ToArray()));
    }

    private static VisioDocument LoadXml(string xml) => LoadXml(Encoding.UTF8.GetBytes(xml));
    private static VisioDocument LoadXml(byte[] bytes) => VisioDocument.LoadLegacyXml(new MemoryStream(bytes)).Value;

    private static string ShapeXml(string flag) =>
        "<VisioDocument xmlns='" + Legacy + "'><Pages><Page ID='0' Name='Flags'><PageSheet><PageProps><PageWidth>4</PageWidth><PageHeight>3</PageHeight></PageProps></PageSheet>" +
        "<Shapes><Shape ID='1' NameU='Controlled' Type='Shape'><XForm><PinX>2</PinX><PinY>1.5</PinY><Width>2</Width><Height>1</Height><LocPinX>1</LocPinX><LocPinY>0.5</LocPinY><Angle>0</Angle></XForm>" +
        "<Line><LineWeight>0.1</LineWeight><LineColor>#0000FF</LineColor><LinePattern>1</LinePattern></Line><Fill><FillForegnd>#FF0000</FillForegnd><FillPattern>1</FillPattern></Fill>" +
        "<Geom IX='0'><NoFill>" + (flag == "NoFill" ? "1" : "0") + "</NoFill><NoLine>" + (flag == "NoLine" ? "1" : "0") + "</NoLine><NoShow>" + (flag == "NoShow" ? "1" : "0") + "</NoShow><NoSnap>0</NoSnap>" +
        "<MoveTo IX='5'><X F='Width*0' Unit='IN'>0</X><Y>0</Y></MoveTo><LineTo IX='9'><X>2</X><Y>0</Y></LineTo><LineTo IX='12'><X>2</X><Y>1</Y></LineTo><LineTo IX='17'><X>0</X><Y>1</Y></LineTo><LineTo IX='21'><X>0</X><Y>0</Y></LineTo></Geom></Shape></Shapes></Page></Pages></VisioDocument>";

    private static string ConnectorXml(string flag) =>
        "<VisioDocument xmlns='" + Legacy + "'><Pages><Page ID='0' Name='Flags'><PageSheet><PageProps><PageWidth>4</PageWidth><PageHeight>3</PageHeight></PageProps></PageSheet>" +
        "<Shapes><Shape ID='1' NameU='Connector' Type='Shape'><XForm><PinX>2</PinX><PinY>1.5</PinY><Width>3</Width><Height>1</Height><LocPinX>1.5</LocPinX><LocPinY>0.5</LocPinY><Angle>0</Angle></XForm>" +
        "<XForm1D><BeginX>0.5</BeginX><BeginY>1</BeginY><EndX>3.5</EndX><EndY>1</EndY></XForm1D><Line><LineWeight>0.1</LineWeight><LineColor>#0000FF</LineColor><LinePattern>1</LinePattern></Line>" +
        "<Geom IX='0'><NoFill>1</NoFill><NoLine>" + (flag == "NoLine" ? "1" : "0") + "</NoLine><NoShow>" + (flag == "NoShow" ? "1" : "0") + "</NoShow><MoveTo IX='5'><X>0</X><Y>0</Y></MoveTo><ArcTo IX='9'><X>3</X><Y>0</Y><A>0.8</A></ArcTo></Geom><Text>X</Text></Shape></Shapes></Page></Pages></VisioDocument>";
}
