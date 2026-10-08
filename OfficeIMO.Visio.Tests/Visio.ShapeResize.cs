using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioShapeResizeTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";

    [Theory]
    [InlineData("direct")]
    [InlineData("xml")]
    [InlineData("openxml")]
    public void ResizingAnEditedProducerGroupPreservesIdentitiesAndAttachedEndpoints(string route) {
        var document = VisioDocument.LoadLegacyXml(Path.Combine(AppContext.BaseDirectory, "Fixtures/LegacyXml/nxbre-ie3.vsx")).Value;
        var master = document.GetMaster("Atom");
        var page = document.AddPage("Resize", 12, 12);
        var instance = page.AddShape("group", master, 4, 4, master.Shape.Width, master.Shape.Height);
        instance.Angle = .4;
        var child = instance.Children[0];
        child.Text = "Edited relation"; child.SetUserCell("Review", "approved");
        child.TextStyle ??= new VisioTextStyle();
        child.TextStyle.TextWidth = child.Width * .75; child.TextStyle.TextHeight = child.Height * .8;
        child.TextStyle.TextPinX = child.Width * .6; child.TextStyle.TextPinY = child.Height * .4;
        child.TextStyle.TextLocPinX = child.Width * .2; child.TextStyle.TextLocPinY = child.Height * .3;
        child.TextStyle.Size = 12; child.TextStyle.LeftMargin = .04;
        var other = page.AddRectangle(9, 4, 1, 1);
        var connector = page.AddConnector(child, other, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
        connector.EndPoint = new OfficePoint(11, 4);
        connector.Waypoints.Add(new VisioConnectorWaypoint(8, 7));
        document = Reopen(document, route); page = document.Pages.Single(p => p.Name == "Resize");
        instance = page.FindShapeById("group")!; child = instance.Children[0]; connector = Assert.Single(page.Connectors);
        VisioConnectionPoint point = connector.FromConnectionPoint!; VisioTextStyle style = child.TextStyle!;
        double originalWidth = instance.Width, originalHeight = instance.Height, childWidth = child.Width, childPin = child.PinX;
        double textWidth = style.TextWidth!.Value, textHeight = style.TextHeight!.Value;
        OfficePoint before = connector.StartPoint;
        XElement sourceMaster = MasterPart(document);

        Assert.Same(instance, page.ResizeShape(instance, originalWidth * 2, originalHeight * .5));
        Assert.Same(child, instance.Children[0]); Assert.Same(style, child.TextStyle); Assert.Same(point, connector.FromConnectionPoint);
        Assert.Equal(childWidth * 2, child.Width, 8); Assert.Equal(childPin * 2, child.PinX, 8);
        Assert.Equal(textWidth * 2, style.TextWidth!.Value, 8); Assert.Equal(textHeight * .5, style.TextHeight!.Value, 8);
        Assert.Equal(12, style.Size); Assert.Equal(.04, style.LeftMargin);
        Assert.Equal(4, instance.PinX); Assert.Equal(4, instance.PinY); Assert.Equal(.4, instance.Angle);
        double cosine = Math.Cos(instance.Angle), sine = Math.Sin(instance.Angle);
        double localX = (before.X - 4) * cosine + (before.Y - 4) * sine;
        double localY = -(before.X - 4) * sine + (before.Y - 4) * cosine;
        Assert.Equal(4 + 2 * localX * cosine - .5 * localY * sine, connector.StartPoint.X, 8);
        Assert.Equal(4 + 2 * localX * sine + .5 * localY * cosine, connector.StartPoint.Y, 8);
        Assert.Equal(new OfficePoint(11, 4), connector.EndPoint);
        Assert.Equal(8, Assert.Single(connector.Waypoints).X); Assert.Equal(7, connector.Waypoints[0].Y);
        Assert.True(XNode.DeepEquals(sourceMaster, MasterPart(document)));
        foreach (var saved in new[] { document, Reopen(document, "xml"), Reopen(document, "openxml") }) {
            var savedPage = saved.Pages.Single(p => p.Name == "Resize");
            var savedChild = savedPage.FindShapeById(child.Id)!;
            Assert.Equal("Edited relation", savedChild.Text); Assert.Equal("approved", savedChild.GetUserCellValue("Review"));
            Assert.Equal(child.Width, savedChild.Width, 8); Assert.Equal(child.PinX, savedChild.PinX, 8);
            Assert.Equal(style.TextWidth!.Value, savedChild.TextStyle!.TextWidth!.Value, 8);
            Assert.Equal(connector.StartPoint.X, savedPage.Connectors[0].StartPoint.X, 8);
            Assert.Equal(connector.StartPoint.Y, savedPage.Connectors[0].StartPoint.Y, 8);
            Assert.Equal(child.Id, savedPage.Connectors[0].From!.Id);
            Assert.Null(savedPage.Connectors[0].To);
            Assert.True(XNode.DeepEquals(sourceMaster, MasterPart(saved)));
            Assert.Contains("Edited relation", savedPage.ToSvg());
        }
    }

    [Fact]
    public void QuarterTurnChildAndTextFramesUseTheirOwnSwappedAxes() {
        var document = VisioDocument.Create(); var page = document.AddPage("Page");
        var group = new VisioShape("group", 4, 4, 4, 4, "") { Type = "Group" };
        var child = new VisioShape("child", 1, 2, 2, 1, "text") { Angle = Math.PI / 2,
            TextStyle = new VisioTextStyle { TextAngle = Math.PI / 2, TextPinX = .3, TextPinY = .4, TextLocPinX = .2, TextLocPinY = .1 } };
        group.Children.Add(child); page.Shapes.Add(group);
        page.ResizeShape(group, 8, 2);
        Assert.Equal(2, child.PinX); Assert.Equal(1, child.PinY);
        Assert.Equal(1, child.Width); Assert.Equal(2, child.Height);
        Assert.Equal(.5, child.LocPinX); Assert.Equal(1, child.LocPinY);
        Assert.Equal(.15, child.TextStyle!.TextPinX!.Value, 8); Assert.Equal(.8, child.TextStyle.TextPinY!.Value, 8);
        Assert.Equal(4, child.TextStyle.TextWidth!.Value); Assert.Equal(.5, child.TextStyle.TextHeight!.Value);
        Assert.Equal(.4, child.TextStyle.TextLocPinX!.Value, 8); Assert.Equal(.05, child.TextStyle.TextLocPinY!.Value, 8);
        foreach (var reopened in new[] { Reopen(document, "xml"), Reopen(document, "openxml") }) {
            var saved = reopened.Pages[0].Shapes[0].Children[0];
            Assert.Equal(1, saved.Width); Assert.Equal(2, saved.Height);
            Assert.Equal(4, saved.TextStyle!.TextWidth); Assert.Equal(.5, saved.TextStyle.TextHeight);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InheritedArcGeometryIsMaterializedThroughNativeResizeAndReopening(bool relative) {
        VisioDocument document = ArcDocument(relative);
        var page = document.Pages.Single(p => p.Name == "Page"); var shape = page.Shapes[0];
        var expected = document.AddPage("Expected", 12, 12).AddShape("expected", document.GetMaster("Arc"), shape.PinX, shape.PinY, 4, 8);
        page.ResizeShape(shape, 4, 8);
        Assert.Equal(PathData(expected.OwnerPage!), PathData(page));
        foreach (var reopened in relative ? new[] { Reopen(document, "openxml") } : new[] { Reopen(document, "xml"), Reopen(document, "openxml") }) {
            var resized = reopened.Pages.Single(p => p.Name == "Page");
            Assert.Equal(PathData(page), PathData(resized));
            using var zip = new ZipArchive(new MemoryStream(reopened.ToBytes()), ZipArchiveMode.Read);
            using var stream = zip.GetEntry("visio/pages/page1.xml")!.Open();
            Assert.Contains(XDocument.Load(stream).Descendants(Modern + "Row"), r => (string?)r.Attribute("T") == (relative ? "RelEllipticalArcTo" : "EllipticalArcTo"));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnsupportedRotatedFramesLeaveTheEntireTreeUnchanged(bool textFrame) {
        var document = VisioDocument.Create(); var page = document.AddPage("Page");
        var root = new VisioShape("group", 4, 4, 4, 4, "") { Type = "Group" };
        root.Children.Add(new VisioShape("first", 1, 1, 1, 1, "valid"));
        var child = new VisioShape("second", 2, 2, 1, 1, "unsupported") { Angle = textFrame ? 0 : .3,
            TextStyle = new VisioTextStyle { TextAngle = textFrame ? .3 : 0, TextWidth = .7 } };
        root.Children.Add(child); page.Shapes.Add(root);
        byte[] before = document.ToLegacyXmlResult().Value;
        Assert.Throws<NotSupportedException>(() => page.ResizeShape(root, 8, 2));
        Assert.Equal(before, document.ToLegacyXmlResult().Value);
        Assert.Same(child, root.Children[1]); Assert.Equal(.7, child.TextStyle!.TextWidth);
        // Uniform scaling supports arbitrary rotations.
        page.ResizeShape(root, 8, 8); Assert.Equal(2, child.Width); Assert.Equal(1.4, child.TextStyle.TextWidth);
    }

    [Fact]
    public void StretchingAnAttachedClosedNativeConnectorRejectsResizeBeforeMutation() {
        const string source = "<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'><Pages><Page ID='0'><Shapes>" +
            "<Shape ID='1' Type='Group'><XForm><PinX>3</PinX><PinY>3</PinY><Width>4</Width><Height>4</Height><LocPinX>0</LocPinX><LocPinY>0</LocPinY></XForm><Shapes><Shape ID='2'><XForm><PinX>1</PinX><PinY>1</PinY><Width>1</Width><Height>1</Height><LocPinX>0</LocPinX><LocPinY>0</LocPinY></XForm></Shape></Shapes></Shape>" +
            "<Shape ID='3'><XForm><PinX>4</PinX><PinY>4</PinY><Width>0</Width><Height>0</Height><LocPinX>0</LocPinX><LocPinY>0</LocPinY></XForm><XForm1D><BeginX>4</BeginX><BeginY>4</BeginY><EndX>4</EndX><EndY>4</EndY></XForm1D><Geom IX='0'><NoFill>1</NoFill><MoveTo IX='1'><X>0</X><Y>0</Y></MoveTo><LineTo IX='2'><X>0</X><Y>0</Y></LineTo></Geom></Shape>" +
            "</Shapes><Connects><Connect FromSheet='3' FromCell='BeginX' ToSheet='2' ToCell='PinX'/></Connects></Page></Pages></VisioDocument>";
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
        var page = document.Pages[0]; var root = page.Shapes[0];
        Assert.Same(root.Children[0], Assert.Single(page.Connectors).From);
        byte[] before = document.ToLegacyXmlResult().Value;
        Assert.Throws<NotSupportedException>(() => page.ResizeShape(root, 8, 4));
        Assert.Equal(before, document.ToLegacyXmlResult().Value);
    }

    [Fact]
    public void ResizeUsesRequestedUnitsAndRejectsInvalidSizesWithoutChangingTheShape() {
        var document = VisioDocument.Create(); var page = document.AddPage("Page"); var shape = page.AddRectangle(4, 4, 1, 1);
        page.ResizeShape(shape, 5.08, 2.54, VisioMeasurementUnit.Centimeters);
        Assert.Equal(2, shape.Width, 8); Assert.Equal(1, shape.Height, 8);
        foreach (double size in new[] { 0, -1, double.NaN, double.PositiveInfinity }) {
            Assert.Throws<ArgumentOutOfRangeException>(() => page.ResizeShape(shape, size, 1));
            Assert.Equal(2, shape.Width, 8); Assert.Equal(1, shape.Height, 8);
        }
        Assert.Throws<InvalidOperationException>(() => page.ResizeShape(new VisioShape("detached"), 1, 1));
    }

    private static string PathData(VisioPage page) => XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions { RenderText = false, RenderStencilArtwork = false })).Descendants()
        .Single(p => p.Name.LocalName == "path" && (string?)p.Attribute("data-officeimo-preserved-geometry") == "true").Attribute("d")!.Value;
    private static VisioDocument Reopen(VisioDocument document, string route) => route == "direct" ? document : route == "xml"
        ? VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value), document.PackageType).Value : VisioDocument.Load(new MemoryStream(document.ToBytes()));
    private static XElement MasterPart(VisioDocument document) {
        using var zip = new ZipArchive(new MemoryStream(document.ToBytes()), ZipArchiveMode.Read);
        var masters = new XElement("Masters");
        foreach (var entry in zip.Entries.Where(e => e.FullName.StartsWith("visio/masters/master", StringComparison.Ordinal) &&
                     e.FullName.EndsWith(".xml", StringComparison.Ordinal) && e.FullName != "visio/masters/masters.xml")
                 .OrderBy(e => e.FullName, StringComparer.Ordinal)) {
            using var stream = entry.Open();
            masters.Add(new XElement("Part", new XAttribute("Name", entry.FullName), XDocument.Load(stream).Root!));
        }
        return masters;
    }
    private static VisioDocument ArcDocument(bool relative) {
        const string source = "<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'><Masters><Master ID='1' NameU='Arc'><Shapes><Shape ID='1'><XForm><Width>4</Width><Height>4</Height><LocPinX>0</LocPinX><LocPinY>0</LocPinY></XForm><Geom IX='0'><NoFill>1</NoFill><MoveTo IX='1'><X>0</X><Y>0</Y></MoveTo><EllipticalArcTo IX='2'><X>3</X><Y>0</Y><A>1.5</A><B>-0.8</B><C>0</C><D>1</D></EllipticalArcTo></Geom></Shape></Shapes></Master></Masters><Pages><Page ID='0' Name='Page'><PageSheet><PageProps><PageWidth>12</PageWidth><PageHeight>12</PageHeight></PageProps></PageSheet><Shapes><Shape ID='2' Master='1'><XForm><PinX>4</PinX><PinY>4</PinY><Width>8</Width><Height>2</Height><LocPinX>0</LocPinX><LocPinY>0</LocPinY></XForm></Shape></Shapes></Page></Pages></VisioDocument>";
        var document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
        if (!relative) return document;
        using var bytes = new MemoryStream(); bytes.Write(document.ToBytes());
        using (var zip = new ZipArchive(bytes, ZipArchiveMode.Update, true)) {
            var entry = zip.GetEntry("visio/masters/master1.xml")!; XDocument xml; using (var stream = entry.Open()) xml = XDocument.Load(stream);
            foreach (var row in xml.Descendants(Modern + "Row")) {
                row.SetAttributeValue("T", "Rel" + (string)row.Attribute("T")!);
                foreach (var cell in row.Elements(Modern + "Cell").Where(c => (string?)c.Attribute("N") is "X" or "Y" or "A" or "B"))
                    cell.SetAttributeValue("V", ((double)cell.Attribute("V")! / 4).ToString("R", CultureInfo.InvariantCulture));
            }
            entry.Delete(); using var output = zip.CreateEntry("visio/masters/master1.xml").Open(); xml.Save(output);
        }
        bytes.Position = 0; return VisioDocument.Load(bytes);
    }
}
