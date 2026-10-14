using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioDrawingScaleRenderingTests {
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private const double Density = 48;

    [Theory]
    [InlineData(0.25, false)]
    [InlineData(2, false)]
    [InlineData(0.5, true)]
    [InlineData(1, false)]
    public void DrawingRatiosProjectPageGeometryAndKeepPhysicalLineWeight(double ratio, bool mixedUnits) {
        VisioDocument document = VisioDocument.Create();
        VisioPage page = document.AddPage("Drawing").Size(12, 8);
        SetScale(page, ratio, mixedUnits);
        VisioShape shape = page.AddRectangle(6, 4, 4, 2, string.Empty);
        shape.FillColor = OfficeColor.Red;
        shape.LineColor = OfficeColor.Blue;
        shape.LineWeight = 0.1;
        document.DuplicatePage(page, "Copy");
        foreach (VisioDocument candidate in Reopened(document)) {
            foreach (VisioPage current in candidate.Pages) {
                byte[] before = candidate.ToLegacyXmlResult().Value;
                XDocument svg = RenderSvg(current);
                Assert.Equal(Math.Ceiling(12 * ratio * Density), Read(svg.Root!, "width"));
                Assert.Equal(Math.Ceiling(8 * ratio * Density), Read(svg.Root!, "height"));
                XElement outline = Assert.Single(svg.Descendants(Svg + "path"));
                double[] coordinates = System.Text.RegularExpressions.Regex.Matches(outline.Attribute("d")!.Value, @"-?\d+(?:\.\d+)?")
                    .Cast<System.Text.RegularExpressions.Match>().Select(match => Parse(match.Value)).ToArray();
                Assert.Equal(4 * ratio * Density, coordinates.Where((_, index) => index % 2 == 0).Min(), 3);
                Assert.Equal(8 * ratio * Density, coordinates.Where((_, index) => index % 2 == 0).Max(), 3);
                Assert.Equal(3 * ratio * Density, coordinates.Where((_, index) => index % 2 == 1).Min(), 3);
                Assert.Equal(5 * ratio * Density, coordinates.Where((_, index) => index % 2 == 1).Max(), 3);
                Assert.Equal(Density * 0.1, Read(outline, "stroke-width"), 3);
                OfficeRasterImage png = RenderPng(current);
                Assert.Equal((int)Math.Ceiling(12 * ratio * Density), png.Width);
                Assert.Equal((int)Math.Ceiling(8 * ratio * Density), png.Height);
                Assert.Equal(OfficeColor.Red, png.GetPixel(png.Width / 2, png.Height / 2));
                int bluePixels = Enumerable.Range(0, png.Width).Count(column => png.GetPixel(column, png.Height / 2) == OfficeColor.Blue);
                Assert.InRange(bluePixels, 8, 12); // Two edges, each approximately .1 physical inch thick.
                Assert.Equal(before, candidate.ToLegacyXmlResult().Value);
                Assert.Equal(12, current.Width);
                Assert.Equal(4, current.Shapes[0].Width);
            }
        }
    }

    [Theory]
    [InlineData(0.25, false)]
    [InlineData(0.25, true)]
    [InlineData(2, false)]
    [InlineData(2, true)]
    public void PlainAndRichFramesRotatePhysicalMarginsWithoutScalingFontSize(double ratio, bool rich) {
        VisioDocument document = TextDocument(ratio, rich);
        foreach (VisioDocument candidate in Reopened(document)) {
            VisioPage page = candidate.Pages[0];
            byte[] before = candidate.ToLegacyXmlResult().Value;
            XDocument svg = RenderSvg(page);
            XElement background = Assert.Single(svg.Descendants(Svg + "rect"), node => node.Attribute("data-officeimo-text-background") != null);
            Assert.Equal(6.1 * Density, Read(background, "x") + Read(background, "width") / 2, 3);
            Assert.Equal((8 - 4.2) * Density, Read(background, "y") + Read(background, "height") / 2, 3);
            Assert.Equal(rich, svg.Descendants().Any(node => node.Attribute("data-officeimo-rich-text") != null));
            Assert.All(svg.Descendants().Attributes("font-size"), font => Assert.Equal(8, Parse(font.Value), 3));
            AssertRedCenter(RenderPng(page), 6.1 * Density, (8 - 4.2) * Density);
            Assert.Equal(before, candidate.ToLegacyXmlResult().Value);
            Assert.Equal(8 / ratio, page.Shapes[0].TextStyle!.TextWidth);
            Assert.Equal(0.4, page.Shapes[0].TextStyle!.LeftMargin);
        }
    }

    [Fact]
    public void RichParagraphIndentsSpacingAndBulletMetricsRemainPhysical() {
        VisioDocument reference = TextDocument(1, rich: true, paragraphs: true);
        VisioDocument drawing = TextDocument(0.25, rich: true, paragraphs: true);
        string[] Expected(XDocument svg) => svg.Descendants(Svg + "text").Select(node => node.ToString()).ToArray();
        Assert.Equal(Expected(RenderSvg(reference.Pages[0])), Expected(RenderSvg(drawing.Pages[0])));
        Assert.Equal(reference.Pages[0].ToPng(PngOptions()), drawing.Pages[0].ToPng(PngOptions()));
    }

    [Fact]
    public void GroupRotationProjectsInheritedFrameAndPhysicalMarginsAfterCopyAndReopening() {
        VisioDocument source = TextDocument(0.25, rich: true);
        XDocument xml = XDocument.Load(new MemoryStream(source.ToLegacyXmlResult().Value));
        XElement shapes = xml.Descendants(Legacy + "Page").Single().Element(Legacy + "Shapes")!;
        XElement child = shapes.Element(Legacy + "Shape")!;
        child.Remove();
        shapes.Add(new XElement(Legacy + "Shape", new XAttribute("ID", "2"), new XAttribute("Type", "Group"),
            new XElement(Legacy + "XForm", new XElement(Legacy + "PinX", "32"), new XElement(Legacy + "PinY", "0"),
                new XElement(Legacy + "LocPinX", "0"), new XElement(Legacy + "LocPinY", "0"),
                new XElement(Legacy + "Width", "48"), new XElement(Legacy + "Height", "32"),
                new XElement(Legacy + "Angle", Math.PI / 2)),
            new XElement(Legacy + "Fill", new XElement(Legacy + "FillPattern", "0")),
            new XElement(Legacy + "Line", new XElement(Legacy + "LinePattern", "0")),
            new XElement(Legacy + "Shapes", child)));
        VisioDocument document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml.ToString()))).Value;
        document.DuplicatePage(document.Pages[0], "Copy");
        foreach (VisioDocument candidate in Reopened(document)) {
            foreach (VisioPage page in candidate.Pages) {
                byte[] before = candidate.ToLegacyXmlResult().Value;
                XElement background = Assert.Single(RenderSvg(page).Descendants(Svg + "rect"), node => node.Attribute("data-officeimo-text-background") != null);
                Assert.Equal(3.8 * Density, Read(background, "x") + Read(background, "width") / 2, 3);
                Assert.Equal((8 - 6.1) * Density, Read(background, "y") + Read(background, "height") / 2, 3);
                AssertRedCenter(RenderPng(page), 3.8 * Density, (8 - 6.1) * Density);
                Assert.Equal(before, candidate.ToLegacyXmlResult().Value);
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeConnectorLabelsAndOverlapSearchUseTheSamePhysicalWorkspace(bool overlap) {
        VisioDocument Create(double ratio) {
            VisioDocument document = VisioDocument.Create();
            VisioPage page = document.AddPage("Labels").Size(8 / ratio, 6 / ratio);
            SetScale(page, ratio);
            VisioConnector connector = page.AddConnector("route", new OfficePoint(1 / ratio, 3 / ratio), new OfficePoint(7 / ratio, 3 / ratio));
            connector.Label = "Review";
            connector.PlaceLabel(0.5, width: 1.4 / ratio, height: 0.4 / ratio);
            connector.LineWeight = 0.08;
            connector.LineColor = OfficeColor.Blue;
            connector.TextStyle = new VisioTextStyle { Size = 12, BackgroundColor = OfficeColor.Red };
            page.AddRectangle(4 / ratio, 3 / ratio, 0.5 / ratio, 0.5 / ratio, string.Empty);
            VisioConnector fallback = page.AddConnector("fallback", new OfficePoint(1 / ratio, 1 / ratio), new OfficePoint(7 / ratio, 1 / ratio));
            fallback.Label = "Fallback";
            fallback.TextStyle = new VisioTextStyle { Size = 12 };
            return document;
        }
        VisioDocument reference = Create(1), drawing = Create(0.25);
        string[] Labels(XDocument svg) => svg.Descendants(Svg + "g")
            .Where(node => node.Attribute("data-visio-connector-id") != null).Select(node => node.ToString()).ToArray();
        Assert.Equal(Labels(RenderSvg(reference.Pages[0], overlap)), Labels(RenderSvg(drawing.Pages[0], overlap)));
        Assert.Equal(reference.Pages[0].ToPng(PngOptions(overlap)), drawing.Pages[0].ToPng(PngOptions(overlap)));
        VisioDocument reopened = VisioDocument.Load(new MemoryStream(drawing.ToBytes()));
        Assert.Equal(5.6, reopened.Pages[0].Connectors[0].LabelPlacement!.Width, 8);
        Assert.Equal(16, reopened.Pages[0].Shapes[0].PinX, 8);
    }

    [Fact]
    public void ImageExportMetadataTargetsAndPixelBudgetUsePhysicalPageDimensions() {
        VisioDocument document = VisioDocument.Create();
        VisioPage page = document.AddPage("Plan").Size(96, 48);
        SetScale(page, 0.125);
        var options = new VisioImageExportOptions { Scale = 0.5, Supersampling = 1, MaximumRasterPixels = 200_000 };
        foreach (OfficeImageExportFormat format in new[] { OfficeImageExportFormat.Svg, OfficeImageExportFormat.Png }) {
            OfficeImageExportResult result = page.ExportImage(format, options);
            Assert.Equal(576, result.Width);
            Assert.Equal(288, result.Height);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.RasterScaleReduced);
            OfficeImageExportResult bounded = page.ExportImage(format, new VisioImageExportOptions { MaximumOutputWidth = 240, Supersampling = 1 });
            Assert.Equal(240, bounded.Width);
            Assert.Equal(120, bounded.Height);
        }
    }

    [Fact]
    public void SupportedStencilArtworkKeepsItsPhysicalPlacementAndProportions() {
        VisioPage Create(double ratio) {
            VisioPage page = VisioDocument.Create().AddPage("Artwork").Size(4 / ratio, 3 / ratio);
            SetScale(page, ratio);
            VisioShape shape = page.AddRectangle(2 / ratio, 1.5 / ratio, 2 / ratio, 1 / ratio, string.Empty);
            shape.SetUserCell("OfficeIMO.StencilId", "event.bus", "STR");
            shape.FillPattern = 0; shape.LinePattern = 0;
            return page;
        }
        VisioPage reference = Create(1), drawing = Create(0.25);
        XElement Artwork(XDocument svg) => svg.Descendants(Svg + "g").Single(node => node.Attribute("data-officeimo-stencil-artwork") != null);
        Assert.Equal(Artwork(RenderSvg(reference)).ToString(), Artwork(RenderSvg(drawing)).ToString());
        Assert.Equal(reference.ToPng(PngOptions()), drawing.ToPng(PngOptions()));
    }

    [Fact]
    public void OriginalPronomPageProjectsToItsIndependentPhysicalDimensionsBeforeRasterAllocation() {
        string fixture = Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", "pronom-visio2002.vtx");
        VisioDocument document = VisioDocument.LoadLegacyXml(fixture).Value;
        XDocument svg = RenderSvg(document.Pages[0]);
        Assert.Equal(36 * Density, Read(svg.Root!, "width"));
        Assert.Equal(24 * Density, Read(svg.Root!, "height"));
        // Allocate only after the SVG has proved the bounded physical page dimensions.
        OfficeRasterImage png = RenderPng(document.Pages[0]);
        Assert.Equal(1728, png.Width);
        Assert.Equal(1152, png.Height);
    }

    private static void SetScale(VisioPage page, double ratio, bool mixed = false) {
        page.PageScale = mixed ? new VisioScaleSetting(2.54, VisioMeasurementUnit.Centimeters) : new VisioScaleSetting(ratio, VisioMeasurementUnit.Inches);
        page.DrawingScale = new VisioScaleSetting(mixed ? 2 : 1, VisioMeasurementUnit.Inches);
    }

    private static VisioDocument TextDocument(double ratio, bool rich, bool paragraphs = false) {
        string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);
        string xml = "<VisioDocument xmlns='" + Legacy + "'><FaceNames><FaceName ID='0' Name='Arial'/></FaceNames><Pages><Page ID='0' Name='Text'>"
            + "<PageSheet><PageProps><PageWidth>" + Number(12 / ratio) + "</PageWidth><PageHeight>" + Number(8 / ratio) + "</PageHeight></PageProps></PageSheet>"
            + "<Shapes><Shape ID='1'><XForm><PinX>" + Number(6 / ratio) + "</PinX><PinY>" + Number(4 / ratio)
            + "</PinY><Width>" + Number(8 / ratio) + "</Width><Height>" + Number(4 / ratio) + "</Height></XForm>"
            + "<Line><LinePattern>0</LinePattern></Line><Fill><FillPattern>0</FillPattern></Fill>"
            + "<TextXForm><TxtPinX>" + Number(4 / ratio) + "</TxtPinX><TxtPinY>" + Number(2 / ratio)
            + "</TxtPinY><TxtWidth>" + Number(8 / ratio) + "</TxtWidth><TxtHeight>" + Number(4 / ratio)
            + "</TxtHeight><TxtLocPinX>" + Number(4 / ratio) + "</TxtLocPinX><TxtLocPinY>" + Number(2 / ratio)
            + "</TxtLocPinY><TxtAngle>" + Number(paragraphs ? 0 : Math.PI / 2) + "</TxtAngle></TextXForm>"
            + "<TextBlock><LeftMargin>0.4</LeftMargin><RightMargin>0</RightMargin><TopMargin>0.2</TopMargin><BottomMargin>0</BottomMargin><VerticalAlign>1</VerticalAlign><TextBkgnd>#FF0000</TextBkgnd></TextBlock>"
            + "<Char IX='0'><Font>0</Font><Color>#000000</Color><Size>0.16666666666666666</Size></Char>"
            + (rich ? "<Char IX='1'><Font>0</Font><Color>#000000</Color><Style>1</Style><Size>0.16666666666666666</Size></Char>" : "")
            + "<Para IX='0'><HorzAlign>1</HorzAlign>" + (paragraphs ? "<IndLeft>0.2</IndLeft><IndFirst>0.1</IndFirst><IndRight>0.2</IndRight><SpBefore>0.1</SpBefore><SpAfter>0.1</SpAfter><SpLine>0.3</SpLine><Bullet>1</Bullet><BulletFontSize>0.16666666666666666</BulletFontSize><TextPosAfterBullet>0.25</TextPosAfterBullet>" : "")
            + "</Para><Text>" + (rich ? "<cp IX='0'/>Hello <cp IX='1'/>world" : "Hello world")
            + "</Text></Shape></Shapes></Page></Pages></VisioDocument>";
        VisioDocument document = VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml))).Value;
        SetScale(document.Pages[0], ratio);
        return document;
    }

    private static IEnumerable<VisioDocument> Reopened(VisioDocument document) {
        yield return document;
        yield return VisioDocument.Load(new MemoryStream(document.ToBytes()));
        yield return VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value;
    }

    private static XDocument RenderSvg(VisioPage page, bool overlap = true) => XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions {
        PixelsPerInch = Density, ResolveConnectorLabelOverlaps = overlap
    }));
    private static VisioPngSaveOptions PngOptions(bool overlap = true) => new() {
        PixelsPerInch = Density, Supersampling = 1, ResolveConnectorLabelOverlaps = overlap
    };
    private static OfficeRasterImage RenderPng(VisioPage page) => VisualBaselineTestSupport.DecodePng(page.ToPng(PngOptions()), "Drawing-scale PNG did not decode.");
    private static double Read(XElement element, string name) => Parse(element.Attribute(name)!.Value);
    private static double Parse(string value) => double.Parse(value, CultureInfo.InvariantCulture);
    private static void AssertRedCenter(OfficeRasterImage image, double x, double y) {
        int left = image.Width, top = image.Height, right = -1, bottom = -1;
        for (int row = 0; row < image.Height; row++)
            for (int column = 0; column < image.Width; column++) {
                OfficeColor color = image.GetPixel(column, row);
                if (color.R > 200 && color.G < 30 && color.B < 30) {
                    left = Math.Min(left, column); right = Math.Max(right, column);
                    top = Math.Min(top, row); bottom = Math.Max(bottom, row);
                }
            }
        Assert.True(right >= left && bottom >= top);
        Assert.InRange((left + right) / 2D, x - 1, x + 1);
        Assert.InRange((top + bottom) / 2D, y - 1, y + 1);
    }
}
