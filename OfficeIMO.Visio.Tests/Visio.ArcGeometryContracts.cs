using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioArcGeometryContractTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";
    private const double Density = 40;

    // Visio ArcTo's signed bow is to the right of the directed chord in its Y-up
    // coordinates. These midpoint controls include both minor and major arcs.
    [Theory]
    [InlineData(0, 0, 3, 0, .8, 1.5, -.8)]
    [InlineData(0, 0, 3, 0, -.8, 1.5, .8)]
    [InlineData(3, 0, 0, 0, .8, 1.5, .8)]
    [InlineData(3, 0, 0, 0, -.8, 1.5, -.8)]
    [InlineData(0, 0, 3, 4, 1, 2.3, 1.4)]
    [InlineData(0, 0, 3, 4, -1, .7, 2.6)]
    [InlineData(0, 0, 3, 0, 2, 1.5, -2)]
    [InlineData(0, 0, 3, 0, -2, 1.5, 2)]
    public void CircularArcShapeAndConnectorPreviewsFollowVisioSignedBow(
        double startX, double startY, double endX, double endY, double bow, double middleX, double middleY) {
        foreach (bool connector in new[] { false, true }) {
            VisioDocument document = LoadArc(startX, startY, endX, endY, bow, connector);
            byte[] before = document.ToLegacyXmlResult().Value;
            VisioPage page = document.Pages[0];
            List<(double X, double Y)> points = PathPoints(GeometryPath(page, connector));
            AssertPoint(points[0], 2 + startX, 4 + startY);
            AssertPoint(points[points.Count - 1], 2 + endX, 4 + endY);
            Assert.Contains(points, point => Near(point, 2 + middleX, 4 + middleY));
            if (Math.Abs(bow) > 1.5 && endY == startY) {
                Assert.True(points.Min(point => point.X) < 2);
                Assert.True(points.Max(point => point.X) > 5);
            }

            OfficeRasterImage image = Raster(page);
            Assert.True(BlueNear(image, 2 + middleX, 4 + middleY), "The PNG must paint the signed arc midpoint.");
            double reflectedX = startX + endX - middleX;
            double reflectedY = startY + endY - middleY;
            Assert.False(BlueNear(image, 2 + reflectedX, 4 + reflectedY), "The PNG must leave the reflected, wrong bow unpainted.");
            Assert.Equal(before, document.ToLegacyXmlResult().Value);
        }
    }

    [Theory]
    [InlineData(.8)]
    [InlineData(-.8)]
    public void CircularArcGeometryAndConnectorLabelsAndArrowsSurviveReopeningAndCopies(double bow) {
        foreach (bool connector in new[] { false, true }) {
            VisioDocument source = LoadArc(0, 0, 3, 0, bow, connector);
            if (connector) {
                source.Pages[0].Connectors[0].Label = "Arc";
                source.Pages[0].Connectors[0].PlaceLabel(.5, width: .4, height: .2);
                source.Pages[0].Connectors[0].EndArrow = EndArrow.Triangle;
            }
            XElement originalGeometry = SavedGeometry(source).Single();
            foreach (VisioDocument document in new[] {
                source,
                LoadXml(source.ToLegacyXmlResult().Value),
                VisioDocument.Load(new MemoryStream(source.ToBytes()))
            }) {
                document.DuplicatePage(document.Pages[0], "Copy");
                byte[] before = document.ToLegacyXmlResult().Value;
                Assert.All(SavedGeometry(document), geometry => Assert.True(XNode.DeepEquals(originalGeometry, geometry)));
                for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
                    VisioPage page = document.Pages[pageIndex];
                    XDocument svg = RenderSvg(page);
                    XElement geometry = GeometryPath(svg, connector);
                    Assert.Contains(PathPoints(geometry), point => Near(point, 3.5, 4 - bow));
                    Assert.True(BlueNear(Raster(page), 3.5, 4 - bow));
                    if (connector) {
                        var snapshot = Assert.Single(document.CreateInspectionSnapshot().Pages[pageIndex].Connectors);
                        Assert.Equal(3.5, snapshot.LabelResolvedPinX!.Value, 8);
                        Assert.Equal(4 - bow, snapshot.LabelResolvedPinY!.Value, 8);
                        XElement label = Assert.Single(svg.Descendants(Svg + "rect"),
                            element => (string?)element.Attribute("data-officeimo-connector-label-background") == "true");
                        Assert.InRange(Number(label, "y") + Number(label, "height") / 2, (6 + bow) * Density - .001, (6 + bow) * Density + .001);
                        XElement arrow = Assert.Single(svg.Descendants(Svg + "path"),
                            element => (string?)element.Attribute("data-officeimo-connector-arrow") == "end");
                        List<(double X, double Y)> arrowPoints = PathPoints(arrow);
                        AssertPoint(arrowPoints[0], 5, 4);
                        double baseY = (arrowPoints[1].Y + arrowPoints[2].Y) / 2;
                        Assert.True(bow > 0 ? baseY < 4 : baseY > 4, "The end arrow must point along the arriving arc tangent.");
                    }
                }
                Assert.Equal(before, document.ToLegacyXmlResult().Value);
            }
        }
    }

    private static XDocument RenderSvg(VisioPage page) => XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions {
        PixelsPerInch = Density, RenderStencilArtwork = false, ResolveConnectorLabelOverlaps = false
    }));

    private static XElement GeometryPath(VisioPage page, bool connector) => GeometryPath(RenderSvg(page), connector);

    private static XElement GeometryPath(XDocument svg, bool connector) => connector
        ? Assert.Single(svg.Descendants(Svg + "g"), element => element.Attribute("data-visio-connector-id") != null).Elements(Svg + "path").First()
        : Assert.Single(svg.Descendants(Svg + "path"), element => (string?)element.Attribute("data-officeimo-preserved-geometry") == "true");

    private static List<(double X, double Y)> PathPoints(XElement path) {
        string[] parts = path.Attribute("d")!.Value.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
        List<(double X, double Y)> points = new();
        for (int i = 0; i < parts.Length; i++) {
            if (parts[i] != "M" && parts[i] != "L") continue;
            double x = double.Parse(parts[++i], CultureInfo.InvariantCulture) / Density;
            double y = 10 - double.Parse(parts[++i], CultureInfo.InvariantCulture) / Density;
            points.Add((x, y));
        }
        return points;
    }

    private static double Number(XElement element, string name) => double.Parse(element.Attribute(name)!.Value, CultureInfo.InvariantCulture);
    private static bool Near((double X, double Y) point, double x, double y) => Math.Abs(point.X - x) < .0001 && Math.Abs(point.Y - y) < .0001;
    private static void AssertPoint((double X, double Y) point, double x, double y) => Assert.True(Near(point, x, y), $"Expected ({x}, {y}), got {point}.");

    private static OfficeRasterImage Raster(VisioPage page) {
        byte[] png = page.ToPng(new VisioPngSaveOptions {
            PixelsPerInch = Density, Supersampling = 1, RenderText = false, RenderConnectorLabels = false,
            RenderStencilArtwork = false, ResolveConnectorLabelOverlaps = false
        });
        Assert.True(OfficeRasterImageDecoder.TryDecode(png, out OfficeRasterImage? image));
        Assert.NotNull(image);
        return image!;
    }

    private static bool BlueNear(OfficeRasterImage image, double x, double y) {
        int px = (int)Math.Round(x * Density), py = (int)Math.Round((10 - y) * Density);
        for (int offsetY = -2; offsetY <= 2; offsetY++) for (int offsetX = -2; offsetX <= 2; offsetX++) {
            OfficeColor color = image.GetPixel(px + offsetX, py + offsetY);
            if (color.B > 200 && color.R < 30 && color.G < 30) return true;
        }
        return false;
    }

    private static XElement[] SavedGeometry(VisioDocument document) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value)).Descendants(Legacy + "Geom").ToArray();
    private static VisioDocument LoadXml(byte[] xml) => VisioDocument.LoadLegacyXml(new MemoryStream(xml)).Value;

    private static VisioDocument LoadArc(double startX, double startY, double endX, double endY, double bow, bool connector) {
        XElement shape = new(Legacy + "Shape", new XAttribute("ID", "1"), new XAttribute("NameU", "NativeArc"), new XAttribute("Type", "Shape"),
            new XElement(Legacy + "XForm", Cell("PinX", 3.5), Cell("PinY", 6), Cell("Width", 3), Cell("Height", 4), Cell("LocPinX", 1.5), Cell("LocPinY", 2), Cell("Angle", 0)),
            new XElement(Legacy + "Line", Cell("LineWeight", .1), new XElement(Legacy + "LineColor", "#0000FF"), Cell("LinePattern", 1)),
            new XElement(Legacy + "Geom", new XAttribute("IX", "0"), Cell("NoFill", 1), Cell("NoLine", 0), Cell("NoShow", 0),
                new XElement(Legacy + "MoveTo", new XAttribute("IX", "5"), Cell("X", startX), Cell("Y", startY)),
                new XElement(Legacy + "ArcTo", new XAttribute("IX", "9"), Cell("X", endX), Cell("Y", endY),
                    new XElement(Legacy + "A", new XAttribute("F", "Width*" + (bow / 3).ToString("R", CultureInfo.InvariantCulture)), new XAttribute("Unit", "IN"), bow.ToString("R", CultureInfo.InvariantCulture)))));
        if (connector) shape.Add(new XElement(Legacy + "XForm1D", Cell("BeginX", 2 + startX), Cell("BeginY", 4 + startY), Cell("EndX", 2 + endX), Cell("EndY", 4 + endY)));
        XDocument xml = new(new XElement(Legacy + "VisioDocument", new XElement(Legacy + "Pages",
            new XElement(Legacy + "Page", new XAttribute("ID", "0"), new XAttribute("Name", "Arc"),
                new XElement(Legacy + "PageSheet", new XElement(Legacy + "PageProps", Cell("PageWidth", 8), Cell("PageHeight", 10))),
                new XElement(Legacy + "Shapes", shape)))));
        return LoadXml(Encoding.UTF8.GetBytes(xml.ToString()));
    }

    private static XElement Cell(string name, double value) => new(Legacy + name, value.ToString("R", CultureInfo.InvariantCulture));
}
