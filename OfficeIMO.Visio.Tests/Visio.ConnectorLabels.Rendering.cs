using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioConnectorLabelRenderingTests {
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ShortCenteredLabelKeepsItsAnchorAfterNativeRoundTrip(bool package, bool resolveOverlaps) {
        VisioDocument document = CreateLabelDocument();
        AssertRenderedCenter(document.Pages[0], 2, 1.5, resolveOverlaps);

        VisioDocument reopened = package
            ? VisioDocument.Load(new MemoryStream(document.ToBytes()))
            : VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value;

        VisioConnector connector = Assert.Single(reopened.Pages[0].Connectors);
        Assert.Equal(0.4, connector.LabelPlacement!.Width, 8);
        Assert.Equal(0.12, connector.LabelPlacement.Height, 8);
        Assert.Equal(2, connector.LabelPlacement.PinX!.Value, 8);
        Assert.Equal(1.5, connector.LabelPlacement.PinY!.Value, 8);
        AssertRenderedCenter(reopened.Pages[0], 2, 1.5, resolveOverlaps);
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(0, true)]
    [InlineData(90, false)]
    [InlineData(90, true)]
    public void NativeOffCenterTextPinRotatesWithinTheAuthoredLabelBox(double angleDegrees, bool resolveOverlaps) {
        XDocument native = XDocument.Load(new MemoryStream(CreateLabelDocument().ToLegacyXmlResult().Value));
        XElement textTransform = Assert.Single(native.Descendants(Legacy + "TextXForm"));
        textTransform.Element(Legacy + "TxtLocPinX")!.Value = "0.05";
        textTransform.Element(Legacy + "TxtLocPinY")!.Value = "0.01";
        textTransform.Element(Legacy + "TxtAngle")!.Value =
            (angleDegrees * Math.PI / 180).ToString("R", CultureInfo.InvariantCulture);
        using var stream = new MemoryStream();
        native.Save(stream);
        stream.Position = 0;
        VisioPage page = VisioDocument.LoadLegacyXml(stream).Value.Pages[0];

        // The authored .4 x .12 box has a center .15 x .05 from its explicit local pin.
        // Renderer minimum fitting dimensions must not become part of that native transform.
        double centerX = angleDegrees == 0 ? 2.15 : 1.95;
        double centerY = angleDegrees == 0 ? 1.55 : 1.65;
        AssertRenderedCenter(page, centerX, centerY, resolveOverlaps);
    }

    [Fact]
    public void OverlapSearchCentersItsExpandedBoundsOnTheAuthoredLabelCenter() {
        VisioDocument document = CreateLabelDocument();
        VisioPage page = document.Pages[0];
        VisioConnector connector = Assert.Single(page.Connectors);
        connector.Label = "Review!";
        connector.PlaceLabel(0.5, width: 0.4, height: 0.12);
        VisioShape obstacle = page.AddRectangle(1.73, 1.5, 0.08, 0.1, string.Empty);

        XDocument svg = XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions {
            PixelsPerInch = 100,
            ResolveConnectorLabelOverlaps = true
        }));
        XElement background = Assert.Single(svg.Descendants(Svg + "rect")
            .Where(rect => (string?)rect.Attribute("data-officeimo-connector-label-background") == "true"));

        Assert.Equal("true", (string?)background.Attribute("data-officeimo-label-adjusted"));
        double labelLeft = double.Parse(background.Attribute("x")!.Value, CultureInfo.InvariantCulture);
        Assert.True(labelLeft >= (obstacle.PinX + obstacle.Width / 2) * 100,
            "The adjusted label still intersects the obstacle.");
    }

    private static VisioDocument CreateLabelDocument() {
        VisioDocument document = VisioDocument.Create();
        VisioPage page = document.AddPage("Short label").Size(4, 3);
        VisioConnector connector = page.AddConnector("label", new OfficePoint(1, 1.5), new OfficePoint(3, 1.5));
        connector.LinePattern = 0;
        connector.Label = "X";
        connector.PlaceLabelAt(2, 1.5, width: 0.4, height: 0.12);
        connector.TextStyle = new VisioTextStyle {
            Size = 9,
            Color = OfficeColor.Black,
            BackgroundColor = OfficeColor.Red,
            HorizontalAlignment = VisioTextHorizontalAlignment.Center,
            VerticalAlignment = VisioTextVerticalAlignment.Middle
        };
        return document;
    }

    private static void AssertRenderedCenter(VisioPage page, double centerX, double centerY, bool resolveOverlaps) {
        const double pixelsPerInch = 100;
        double expectedX = centerX * pixelsPerInch;
        double expectedY = (page.Height - centerY) * pixelsPerInch;
        XDocument svg = XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions {
            PixelsPerInch = pixelsPerInch,
            ResolveConnectorLabelOverlaps = resolveOverlaps
        }));
        XElement background = Assert.Single(svg.Descendants(Svg + "rect")
            .Where(rect => (string?)rect.Attribute("data-officeimo-connector-label-background") == "true"));
        double Read(string name) => double.Parse(background.Attribute(name)!.Value, CultureInfo.InvariantCulture);
        Assert.Equal(expectedX, Read("x") + Read("width") / 2, 3);
        Assert.Equal(expectedY, Read("y") + Read("height") / 2, 3);

        OfficeRasterImage raster = VisualBaselineTestSupport.DecodePng(page.ToPng(new VisioPngSaveOptions {
            PixelsPerInch = pixelsPerInch,
            Supersampling = 1,
            ResolveConnectorLabelOverlaps = resolveOverlaps
        }), "Connector label PNG could not be decoded.");
        int left = raster.Width, top = raster.Height, right = -1, bottom = -1;
        for (int y = 0; y < raster.Height; y++) {
            for (int x = 0; x < raster.Width; x++) {
                OfficeColor color = raster.GetPixel(x, y);
                if (color.R > 200 && color.G < 30 && color.B < 30) {
                    left = Math.Min(left, x);
                    top = Math.Min(top, y);
                    right = Math.Max(right, x);
                    bottom = Math.Max(bottom, y);
                }
            }
        }
        Assert.True(right >= left && bottom >= top, "Connector label background was not rendered.");
        Assert.InRange((left + right) / 2D, expectedX - 1, expectedX + 1);
        Assert.InRange((top + bottom) / 2D, expectedY - 1, expectedY + 1);
    }
}
