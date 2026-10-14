using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioTextFramePlacementTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";

    [Theory]
    [InlineData(true, false, false, false, 2.9, 2.8, -90)]
    [InlineData(false, true, false, false, 2.9, 2.8, -90)]
    [InlineData(true, true, false, false, 3.1, 3.2, 90)]
    [InlineData(false, false, true, false, 2.9, 4.2, -90)]
    [InlineData(false, false, false, true, 4.1, 2.8, -90)]
    [InlineData(false, false, true, true, 2.9, 2.8, 90)]
    [InlineData(false, true, true, false, 3.1, 3.8, 90)]
    public void ReflectedTextLocalPinsAndMarginsFollowNativeTextThenShapeTransformOrder(
        bool ownX, bool ownY, bool parentX, bool parentY, double x, double y, double angle) {
        foreach (bool rich in new[] { false, true }) {
            string shape = Shape("2", "3", "3", "4", "2", 0,
                Frame(2, 1, 2, 1, 1, 0.5, Math.PI / 2) + Text(rich, left: 0.4, top: 0.2), flipX: ownX, flipY: ownY);
            if (parentX || parentY)
                shape = Shape("1", parentX ? "6" : "1", parentY ? "6" : "1", "6", "6", 0,
                    "<Shapes>" + shape + "</Shapes>", localPinZero: true, flipX: parentX, flipY: parentY);
            foreach (VisioDocument document in Reopened(Load("", shape)))
                AssertPlacement(document.Pages[0], x, y, angle, rich);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RotatedLocalPinAndNestedParentAnglesPlacePlainAndRichTextOnThePage(bool rich) {
        string child = Shape("3", "2", "1", "2", "1", Math.PI / 2,
            Frame(0.75, 0.25, 2, 0.5, 2, 0.25, -Math.PI / 2) + Text(rich));
        string parent = Shape("2", "0", "0", "4", "4", Math.PI / 4, "<Shapes>" + child + "</Shapes>", localPinZero: true);
        string outer = Shape("1", "4", "4", "4", "4", Math.PI / 4, "<Shapes>" + parent + "</Shapes>", localPinZero: true);
        VisioDocument document = Load("", outer);
        foreach (VisioDocument candidate in Reopened(document))
            AssertPlacement(candidate.Pages[0], 3.25, 5.25, 90, rich);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AsymmetricMarginsMoveTheContentCenterInTheRotatedTextFrame(bool rich) {
        string shape = Shape("1", "3", "3", "4", "2", 0,
            Frame(2, 1, 2, 1, 1, 0.5, Math.PI / 2) + Text(rich, left: 0.4, top: 0.2));
        VisioDocument document = Load("", shape);
        foreach (VisioDocument candidate in Reopened(document))
            AssertPlacement(candidate.Pages[0], 3.1, 3.2, 90, rich);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeFrameCenterDoesNotMoveWhenRendererEnlargesTinyFittingBounds(bool rich) {
        string shape = Shape("1", "3", "3", "4", "2", 0,
            Frame(2, 1, 0.02, 0.04, 0, 0, Math.PI / 2) + Text(rich, segment: "X"));
        VisioDocument document = Load("", shape);
        foreach (VisioDocument candidate in Reopened(document))
            AssertPlacement(candidate.Pages[0], 2.98, 3.01, 90, rich);
    }

    [Fact]
    public void CachedMasterFramesScaleForNewInstancesAndLocalInheritedCachesSurviveCopiesAndReopening() {
        string master = "<Master ID='1' NameU='Cached frame'><Shapes>"
            + Shape("1", "1", "0.5", "2", "1", Math.PI / 2,
                Frame(1.5, 0.5, 1, 0.25, 1, 0.125, -Math.PI / 2) + Text(rich: true))
            + "</Shapes></Master>";
        string instance = Shape("1", "3", "3", "2", "1", Math.PI / 2,
            Frame(0.5, 0.5, 1.5, 0.5, 1.5, 0.25, -Math.PI / 2, inherited: true) + Text(rich: false), master: "1");
        VisioDocument document = Load(master, instance);
        VisioPage scaledPage = document.AddPage("Scaled").Size(8, 8);
        VisioShape scaled = scaledPage.AddShape("scaled", document.Masters.Single(), 4, 4, 4, 3);
        Assert.Equal(3, scaled.TextStyle!.TextPinX);
        Assert.Equal(1.5, scaled.TextStyle.TextPinY);
        Assert.Equal(2, scaled.TextStyle.TextWidth);
        Assert.Equal(0.75, scaled.TextStyle.TextHeight);
        Assert.Equal(2, scaled.TextStyle.TextLocPinX);
        Assert.Equal(0.375, scaled.TextStyle.TextLocPinY);
        foreach (VisioPage page in document.Pages.ToArray()) {
            page.SelectShapes(_ => true).Duplicate(new VisioShapeDuplicationOptions { OffsetX = 0, OffsetY = 0 });
            document.DuplicatePage(page, page.Name + " copy");
        }
        foreach (VisioDocument candidate in Reopened(document)) {
            foreach (VisioPage page in candidate.Pages)
                AssertPlacement(page, page.Name!.StartsWith("Scaled", StringComparison.Ordinal) ? 3 : 2.25,
                    page.Name.StartsWith("Scaled", StringComparison.Ordinal) ? 5 : 2.5, 0,
                    rich: page.Name.StartsWith("Scaled", StringComparison.Ordinal) ? true : null, backgrounds: 2);
            XDocument xml = XDocument.Load(new MemoryStream(candidate.ToLegacyXmlResult().Value));
            XElement[] localCaches = xml.Descendants(Legacy + "Page").Where(page => !((string?)page.Attribute("Name") ?? "").StartsWith("Scaled", StringComparison.Ordinal))
                .Descendants(Legacy + "TxtPinX").ToArray();
            Assert.Equal(4, localCaches.Length);
            Assert.All(localCaches, cell => Assert.Equal(0.5, double.Parse(cell.Value, CultureInfo.InvariantCulture)));
        }
    }

    [Fact]
    public void AuthoredSingleShapeMasterSuppliesItsTextStyleAndFrameToInstancesAndNativeMasterContent() {
        VisioDocument document = VisioDocument.Create();
        var blueprint = new VisioShape("1", 1, 0.5, 2, 1, "MMMMMM") {
            Angle = Math.PI / 2, FillPattern = 0, LinePattern = 0,
            TextStyle = new VisioTextStyle {
                FontFamily = "Arial", Size = 16, Color = OfficeColor.Black, BackgroundColor = OfficeColor.Red,
                HorizontalAlignment = VisioTextHorizontalAlignment.Center, VerticalAlignment = VisioTextVerticalAlignment.Middle,
                LeftMargin = 0, RightMargin = 0, TopMargin = 0, BottomMargin = 0,
                TextPinX = 1.5, TextPinY = 0.5, TextWidth = 1, TextHeight = 0.25,
                TextLocPinX = 1, TextLocPinY = 0.125, TextAngle = -Math.PI / 2
            }
        };
        VisioMaster master = document.RegisterMaster("Authored frame", blueprint);
        VisioPage page = document.AddPage("Authored").Size(8, 8);
        VisioShape instance = page.AddShape("instance", master, 3, 3, 2, 1);
        Assert.NotNull(instance.TextStyle);
        Assert.NotSame(blueprint.TextStyle, instance.TextStyle);
        foreach (VisioDocument candidate in Reopened(document)) {
            VisioTextStyle style = candidate.Pages[0].Shapes[0].TextStyle!;
            Assert.Equal("Arial", style.FontFamily);
            Assert.Equal(16, style.Size);
            Assert.Equal(1, style.TextWidth);
            Assert.Equal(-Math.PI / 2, style.TextAngle!.Value, 12);
            AssertPlacement(candidate.Pages[0], 2.5, 3.5, 0, rich: null);
            XDocument native = XDocument.Load(new MemoryStream(candidate.ToLegacyXmlResult().Value));
            XElement frame = Assert.Single(native.Descendants(Legacy + "Master")).Descendants(Legacy + "TextXForm").Single();
            Assert.Equal(1.5, double.Parse(frame.Element(Legacy + "TxtPinX")!.Value, CultureInfo.InvariantCulture));
            Assert.Equal(1, double.Parse(frame.Element(Legacy + "TxtLocPinX")!.Value, CultureInfo.InvariantCulture));
        }
        Assert.Equal(1.5, blueprint.TextStyle!.TextPinX);
        instance.TextStyle!.TextPinX = 1;
        Assert.Equal(1.5, blueprint.TextStyle.TextPinX);
    }

    private static VisioDocument Load(string masters, string shapes) {
        string xml = "<VisioDocument xmlns='" + Legacy + "'><FaceNames><FaceName ID='0' Name='Arial'/></FaceNames>"
            + "<Masters>" + masters + "</Masters><Pages><Page ID='0' Name='Local'><PageSheet><PageProps>"
            + "<PageWidth>8</PageWidth><PageHeight>8</PageHeight></PageProps></PageSheet><Shapes>" + shapes + "</Shapes></Page></Pages></VisioDocument>";
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(xml))).Value;
    }

    private static string Shape(string id, string x, string y, string width, string height, double angle, string body,
        bool localPinZero = false, string? master = null, bool flipX = false, bool flipY = false) => "<Shape ID='" + id + "'" + (master == null ? "" : " Master='" + master + "'")
        + (body.Contains("<Shapes>") ? " Type='Group'" : "") + "><XForm><PinX>" + x + "</PinX><PinY>" + y
        + "</PinY><Width>" + width + "</Width><Height>" + height + "</Height><Angle>" + Number(angle)
        + "</Angle>" + (localPinZero ? "<LocPinX>0</LocPinX><LocPinY>0</LocPinY>" : "")
        + "<FlipX>" + (flipX ? "1" : "0") + "</FlipX><FlipY>" + (flipY ? "1" : "0")
        + "</FlipY></XForm><Line><LinePattern>0</LinePattern></Line><Fill><FillPattern>0</FillPattern></Fill>" + body + "</Shape>";

    private static string Frame(double x, double y, double width, double height, double localX, double localY, double angle, bool inherited = false) {
        string Cell(string name, double value) => "<" + name + (inherited ? " F='Inh'" : "") + ">" + Number(value) + "</" + name + ">";
        return "<TextXForm>" + Cell("TxtPinX", x) + Cell("TxtPinY", y) + Cell("TxtWidth", width) + Cell("TxtHeight", height)
            + Cell("TxtLocPinX", localX) + Cell("TxtLocPinY", localY) + Cell("TxtAngle", angle) + "</TextXForm>";
    }

    private static string Text(bool rich, double left = 0, double top = 0, string segment = "MMM") => "<TextBlock><LeftMargin>" + Number(left)
        + "</LeftMargin><RightMargin>0</RightMargin><TopMargin>" + Number(top)
        + "</TopMargin><BottomMargin>0</BottomMargin><VerticalAlign>1</VerticalAlign><TextBkgnd>#FF0000</TextBkgnd></TextBlock>"
        + "<Char IX='0'><Font>0</Font><Color>#000000</Color><Size>0.2222222222222222</Size></Char>"
        + (rich ? "<Char IX='1'><Font>0</Font><Color>#000000</Color><Style>1</Style><Size>0.2222222222222222</Size></Char>" : "")
        + "<Para IX='0'><HorzAlign>1</HorzAlign></Para><Text>" + (rich ? "<cp IX='0'/>" + segment + "<cp IX='1'/>" + segment : segment + segment) + "</Text>";

    private static string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);

    private static IEnumerable<VisioDocument> Reopened(VisioDocument document) {
        yield return document;
        yield return VisioDocument.Load(new MemoryStream(document.ToBytes()));
        yield return VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value;
    }

    private static void AssertPlacement(VisioPage page, double centerX, double centerY, double degrees, bool? rich, int backgrounds = 1) {
        byte[] before = page.OwnerDocument!.ToLegacyXmlResult().Value;
        const double scale = 100;
        double x = centerX * scale, y = (page.Height - centerY) * scale;
        XDocument svg = XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = scale, RenderStencilArtwork = false }));
        OfficeRasterImage raster = VisualBaselineTestSupport.DecodePng(page.ToPng(new VisioPngSaveOptions {
            PixelsPerInch = scale, Supersampling = 1, RenderStencilArtwork = false
        }), "Text-frame PNG could not be decoded.");
        Assert.Multiple(() => {
            XElement[] rectangles = svg.Descendants(Svg + "rect").Where(rect => rect.Attribute("data-officeimo-text-background") != null).ToArray();
            Assert.Equal(backgrounds, rectangles.Length);
            foreach (XElement rectangle in rectangles) {
                double Read(string name) => double.Parse(rectangle.Attribute(name)!.Value, CultureInfo.InvariantCulture);
                Assert.Equal(x, Read("x") + Read("width") / 2, 3);
                Assert.Equal(y, Read("y") + Read("height") / 2, 3);
                if (degrees == 0) Assert.Null(rectangle.Attribute("transform"));
                else {
                    string[] rotation = rectangle.Attribute("transform")!.Value.Split(new[] { '(', ')', ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
                    Assert.Equal("rotate", rotation[0]);
                    Assert.Equal(4, rotation.Length);
                    Assert.Equal(-degrees, double.Parse(rotation[1], CultureInfo.InvariantCulture), 8);
                    Assert.Equal(x, double.Parse(rotation[2], CultureInfo.InvariantCulture), 3);
                    Assert.Equal(y, double.Parse(rotation[3], CultureInfo.InvariantCulture), 3);
                }
            }
            if (rich.HasValue) Assert.Equal(rich.Value, svg.Descendants().Any(node => node.Attribute("data-officeimo-rich-text") != null));
        }, () => {
            int left = raster.Width, top = raster.Height, right = -1, bottom = -1;
            for (int row = 0; row < raster.Height; row++)
                for (int column = 0; column < raster.Width; column++) {
                    OfficeColor color = raster.GetPixel(column, row);
                    if (color.R > 200 && color.G < 30 && color.B < 30) {
                        left = Math.Min(left, column); right = Math.Max(right, column);
                        top = Math.Min(top, row); bottom = Math.Max(bottom, row);
                    }
                }
            Assert.True(right >= left && bottom >= top, "Text background was not rendered.");
            Assert.InRange((left + right) / 2D, x - 1, x + 1);
            Assert.InRange((top + bottom) / 2D, y - 1, y + 1);
            Assert.True(Math.Abs(degrees) == 90 ? bottom - top > right - left : right - left > bottom - top,
                "Text background orientation does not match the page text frame.");
        });
        Assert.Equal(before, page.OwnerDocument.ToLegacyXmlResult().Value);
    }
}
