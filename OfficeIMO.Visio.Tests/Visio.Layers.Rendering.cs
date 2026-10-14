using System.Globalization;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioLayerRenderingTests {
    private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";
    private const double PixelsPerInch = 40;

    [Fact]
    public void DefaultPreviewUsesScreenLayersAcrossRetainedAndCanonicalApis() {
        VisioDocument document = CreateLayerDocument();
        VisioPage page = document.Pages[0];

        AssertLayerShapes(XDocument.Parse(page.ToSvg()), VisioLayerRenderMode.Visible);
        AssertLayerShapes(XDocument.Parse(document.ToSvg()), VisioLayerRenderMode.Visible);
        AssertLayerShapes(ReadSvg(page.ExportImage(OfficeImageExportFormat.Svg)), VisioLayerRenderMode.Visible);
        AssertLayerPixels(Decode(page.ToPng()), 96, VisioLayerRenderMode.Visible);
        AssertLayerPixels(Decode(document.ToPng()), 96, VisioLayerRenderMode.Visible);
        AssertLayerPixels(Decode(page.ExportImage(OfficeImageExportFormat.Png).Bytes), 96, VisioLayerRenderMode.Visible);
    }

    [Theory]
    [InlineData(VisioLayerRenderMode.Visible)]
    [InlineData(VisioLayerRenderMode.Printable)]
    [InlineData(VisioLayerRenderMode.All)]
    public void LayerModesKeepIndependentFlagsAndMixedMembershipAfterNativeReload(VisioLayerRenderMode mode) {
        foreach (VisioDocument document in NativeVersions(CreateLayerDocument())) {
            VisioPage page = document.Pages[0];
            byte[] nativeBefore = document.ToLegacyXmlResult().Value;
            string svg = page.ToSvg(new VisioSvgSaveOptions { LayerMode = mode });
            AssertLayerShapes(XDocument.Parse(svg), mode);
            Assert.Equal(mode == VisioLayerRenderMode.All, svg.Contains("Disabled text"));
            Assert.Equal(mode != VisioLayerRenderMode.Printable, svg.Contains("Screen text"));
            Assert.Equal(mode != VisioLayerRenderMode.Visible, svg.Contains("Print text"));
            Assert.Contains("Mixed text", svg);
            Assert.Contains("Unassigned text", svg);

            byte[] png = page.ToPng(new VisioPngSaveOptions {
                LayerMode = mode, PixelsPerInch = PixelsPerInch, Supersampling = 1, RenderText = false
            });
            AssertLayerPixels(Decode(png), PixelsPerInch, mode);
            OfficeImageExportResult canonical = document.ExportImage(OfficeImageExportFormat.Png, new VisioImageExportOptions {
                LayerMode = mode, TargetDpi = PixelsPerInch, Supersampling = 1, RenderText = false
            });
            AssertLayerPixels(Decode(canonical.Bytes), PixelsPerInch, mode);

            Assert.Equal(nativeBefore, document.ToLegacyXmlResult().Value);
            Assert.Equal(5, page.Shapes.Count);
            Assert.False(page.FindLayer("Disabled")!.Visible);
            Assert.False(page.FindLayer("Disabled")!.Print);
            Assert.Equal(2, page.Shapes.Single(shape => shape.NameU == "Mixed").LayerNames.Count);
        }
    }

    [Fact]
    public void DisplayAndUniversalNamesUseCaseInsensitiveFirstMatchAndUnknownMembershipRemainsVisible() {
        VisioDocument document = VisioDocument.Create();
        VisioPage page = document.AddPage("Aliases").Size(4, 2);
        page.Layers.Add(new VisioLayer("Translated name", "Shared") { Visible = false, Print = false });
        page.Layers.Add(new VisioLayer("Shared", "Other") { Visible = true, Print = true });
        VisioShape first = AddColoredShape(page, "First match", 0.7, OfficeColor.Red);
        first.LayerNames.Add("sHaReD");
        VisioShape second = AddColoredShape(page, "Universal", 2, OfficeColor.Blue);
        second.LayerNames.Add("oThEr");
        VisioShape unknown = AddColoredShape(page, "Unknown", 3.3, OfficeColor.Green);
        unknown.LayerNames.Add("Translated name");
        unknown.LayerNames.Add("Undeclared layer");

        foreach (VisioLayerRenderMode mode in new[] { VisioLayerRenderMode.Visible, VisioLayerRenderMode.Printable }) {
            XDocument svg = XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions { LayerMode = mode }));
            AssertShape(svg, "First match", false);
            AssertShape(svg, "Universal", true);
            AssertShape(svg, "Unknown", true);
            OfficeRasterImage raster = Decode(page.ToPng(new VisioPngSaveOptions {
                LayerMode = mode, PixelsPerInch = PixelsPerInch, Supersampling = 1, RenderText = false
            }));
            AssertPixel(raster, PixelsPerInch, page.Height, 0.7, 1.3, OfficeColor.White);
            AssertPixel(raster, PixelsPerInch, page.Height, 2, 1.3, OfficeColor.Blue);
            AssertPixel(raster, PixelsPerInch, page.Height, 3.3, 1.3, OfficeColor.Green);
        }
    }

    [Fact]
    public void HiddenGroupKeepsVisibleAndUnassignedChildrenWithTheirNativeTransformsAndDiagnostics() {
        VisioDocument document = VisioDocument.Create();
        VisioPage page = document.AddPage("Independent children").Size(4, 2);
        page.AddLayer("Hidden").Visible = false;
        var group = new VisioShape("1", 2, 1, 4, 2, "Hidden parent") {
            Type = "Group", NameU = "Parent", FillColor = OfficeColor.Red, LinePattern = 0,
            TextStyle = new VisioTextStyle { FontFamily = "OfficeIMO Missing Hidden Parent Font" }
        };
        var visible = new VisioShape("2", 0.7, 1, 0.8, 0.8, "Visible child") {
            NameU = "Visible child", FillColor = OfficeColor.Blue, LinePattern = 0,
            TextStyle = new VisioTextStyle { FontFamily = "OfficeIMO Missing Visible Child Font" }
        };
        var unassigned = new VisioShape("3", 2, 1, 0.8, 0.8, "Unassigned child") {
            NameU = "Unassigned child", FillColor = OfficeColor.Green, LinePattern = 0
        };
        var hidden = new VisioShape("4", 3.3, 1, 0.8, 0.8, "Hidden child") {
            NameU = "Hidden child", FillColor = OfficeColor.Red, LinePattern = 0
        };
        group.Children.Add(visible); group.Children.Add(unassigned); group.Children.Add(hidden);
        page.Shapes.Add(group);
        page.AddToLayer("Hidden", group).AddToLayer("Hidden", hidden);
        page.AddToLayer("Shown", visible);

        foreach (VisioDocument candidate in NativeVersions(document)) {
            VisioPage current = candidate.Pages[0];
            OfficeImageExportResult result = current.ExportImage(OfficeImageExportFormat.Svg);
            XDocument svg = ReadSvg(result);
            AssertShape(svg, "Parent", false);
            AssertShape(svg, "Visible child", true);
            AssertShape(svg, "Unassigned child", true);
            AssertShape(svg, "Hidden child", false);
            Assert.DoesNotContain("Hidden parent", Encoding.UTF8.GetString(result.Bytes));
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Message.Contains("OfficeIMO Missing Hidden Parent Font"));
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Message.Contains("OfficeIMO Missing Visible Child Font"));
            OfficeRasterImage raster = Decode(current.ToPng(new VisioPngSaveOptions {
                PixelsPerInch = PixelsPerInch, Supersampling = 1, RenderText = false
            }));
            AssertPixel(raster, PixelsPerInch, current.Height, 0.7, 1.3, OfficeColor.Blue);
            AssertPixel(raster, PixelsPerInch, current.Height, 2, 1.3, OfficeColor.Green);
            AssertPixel(raster, PixelsPerInch, current.Height, 3.3, 1.3, OfficeColor.White);
            AssertPixel(raster, PixelsPerInch, current.Height, 0.1, 0.1, OfficeColor.White);
        }
    }

    [Theory]
    [InlineData(VisioLayerRenderMode.Visible)]
    [InlineData(VisioLayerRenderMode.Printable)]
    [InlineData(VisioLayerRenderMode.All)]
    public void ConnectorLineAndLabelUseOwnMembershipIndependentlyOfHiddenEndpoints(VisioLayerRenderMode mode) {
        VisioDocument document = VisioDocument.Create();
        VisioPage page = document.AddPage("Independent connectors").Size(5, 3);
        page.AddLayer("Disabled").Visible = false;
        page.FindLayer("Disabled")!.Print = false;
        page.AddLayer("Screen").Print = false;
        VisioShape first = page.AddRectangle(1, 2.5, 0.6, 0.6, string.Empty);
        VisioShape second = page.AddRectangle(4, 2.5, 0.6, 0.6, string.Empty);
        first.FillColor = second.FillColor = OfficeColor.Green;
        first.LinePattern = second.LinePattern = 0;
        page.AddToLayer("Disabled", first).AddToLayer("Disabled", second);
        VisioConnector screen = page.AddConnector(first, second, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
        screen.Label = "ScreenLink";
        screen.LineColor = OfficeColor.Red; screen.LineWeight = 0.08;
        screen.EndArrow = screen.BeginArrow = EndArrow.None;
        page.AddToLayer("Screen", screen);
        VisioConnector hidden = page.AddConnector("hidden", new OfficePoint(1, 0.8), new OfficePoint(4, 0.8));
        hidden.Label = "HiddenLink"; hidden.LineColor = OfficeColor.Blue; hidden.LineWeight = 0.08;
        hidden.EndArrow = hidden.BeginArrow = EndArrow.None;
        page.AddToLayer("Disabled", hidden);

        foreach (VisioDocument candidate in NativeVersions(document)) {
            VisioPage current = candidate.Pages[0];
            string svg = current.ToSvg(new VisioSvgSaveOptions { LayerMode = mode });
            Assert.Equal(mode != VisioLayerRenderMode.Printable, svg.Contains("ScreenLink"));
            Assert.Equal(mode == VisioLayerRenderMode.All, svg.Contains("HiddenLink"));
            OfficeRasterImage raster = Decode(current.ToPng(new VisioPngSaveOptions {
                LayerMode = mode, PixelsPerInch = PixelsPerInch, Supersampling = 1,
                RenderText = false, RenderConnectorLabels = false
            }));
            AssertPixel(raster, PixelsPerInch, current.Height, 2.5, 2.5,
                mode == VisioLayerRenderMode.Printable ? OfficeColor.White : OfficeColor.Red);
            AssertPixel(raster, PixelsPerInch, current.Height, 2.5, 0.8,
                mode == VisioLayerRenderMode.All ? OfficeColor.Blue : OfficeColor.White);
            AssertPixel(raster, PixelsPerInch, current.Height, 1, 2.7,
                mode == VisioLayerRenderMode.All ? OfficeColor.Green : OfficeColor.White);
            Assert.Equal(2, current.Shapes.Count);
            Assert.Equal(2, current.Connectors.Count);
        }
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Svg)]
    [InlineData(OfficeImageExportFormat.Png)]
    public void HiddenTextAndForeignContentDoNotTriggerExportLossPolicies(OfficeImageExportFormat format) {
        VisioDocument document = VisioDocument.Create();
        VisioPage page = document.AddPage("Hidden diagnostics").Size(3, 2);
        page.AddLayer("Disabled").Visible = false;
        page.FindLayer("Disabled")!.Print = false;
        VisioShape text = page.AddRectangle(0.7, 1, 0.8, 0.8, "Hidden text");
        text.TextStyle = new VisioTextStyle { FontFamily = "OfficeIMO Missing Hidden Shape Font" };
        VisioShape foreign = page.AddRectangle(2, 1, 0.8, 0.8, string.Empty);
        foreign.Type = "Foreign";
        VisioConnector connector = page.AddConnector("label", new OfficePoint(0.5, 0.2), new OfficePoint(2.5, 0.2));
        connector.Label = "Hidden label";
        connector.TextStyle = new VisioTextStyle { FontFamily = "OfficeIMO Missing Hidden Connector Font" };
        page.AddToLayer("Disabled", text).AddToLayer("Disabled", foreign).AddToLayer("Disabled", connector);
        var options = new VisioImageExportOptions {
            Supersampling = 1, Policy = new OfficeImageExportPolicy { RequireNoLoss = true }
        };

        foreach (VisioLayerRenderMode mode in new[] { VisioLayerRenderMode.Visible, VisioLayerRenderMode.Printable }) {
            options.LayerMode = mode;
            OfficeImageExportResult result = page.ExportImage(format, options);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.FontSubstituted);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
        }
        options.LayerMode = VisioLayerRenderMode.All;
        OfficeImageExportPolicyException exception = Assert.Throws<OfficeImageExportPolicyException>(() => page.ExportImage(format, options));
        Assert.Contains(exception.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
        Assert.Contains(exception.Diagnostics, diagnostic => diagnostic.Message.Contains("OfficeIMO Missing Hidden Shape Font"));
        Assert.Contains(exception.Diagnostics, diagnostic => diagnostic.Message.Contains("OfficeIMO Missing Hidden Connector Font"));
    }

    [Theory]
    [InlineData("shape")]
    [InlineData("line")]
    [InlineData("label")]
    public void HiddenObstaclesDoNotDisplaceVisibleConnectorLabels(string obstacleKind) {
        VisioDocument document = VisioDocument.Create();
        VisioPage page = document.AddPage("Label visibility").Size(4, 3);
        page.AddLayer("Hidden").Visible = false;
        if (obstacleKind == "shape") {
            VisioShape obstacle = page.AddRectangle(1.73, 1.5, 0.08, 0.1, string.Empty);
            page.AddToLayer("Hidden", obstacle);
        } else {
            VisioConnector obstacle = page.AddConnector("obstacle", new OfficePoint(2, 1), new OfficePoint(2, 2));
            obstacle.LineWeight = 0.1;
            if (obstacleKind == "label") {
                obstacle.LinePattern = 0;
                obstacle.Label = "Hidden label";
                obstacle.PlaceLabelAt(2, 1.5, width: 1.35, height: 0.34);
            }
            page.AddToLayer("Hidden", obstacle);
        }
        VisioConnector connector = page.AddConnector("visible", new OfficePoint(1, 1.5), new OfficePoint(3, 1.5));
        connector.LinePattern = 0; connector.Label = "X";
        connector.PlaceLabel(0.5, width: 0.4, height: 0.12);
        connector.TextStyle = new VisioTextStyle { BackgroundColor = OfficeColor.Red, Color = OfficeColor.Transparent };

        XDocument visibleSvg = XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions { PixelsPerInch = 100 }));
        XElement visibleLabel = LabelBackground(visibleSvg, connector.Id);
        Assert.NotEqual("true", (string?)visibleLabel.Attribute("data-officeimo-label-adjusted"));
        Assert.Equal(200, ReadNumber(visibleLabel, "x") + ReadNumber(visibleLabel, "width") / 2, 3);
        Assert.Equal(150, ReadNumber(visibleLabel, "y") + ReadNumber(visibleLabel, "height") / 2, 3);
        OfficeRasterImage raster = Decode(page.ToPng(new VisioPngSaveOptions { PixelsPerInch = 100, Supersampling = 1 }));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(200, 150));

        XDocument allSvg = XDocument.Parse(page.ToSvg(new VisioSvgSaveOptions { LayerMode = VisioLayerRenderMode.All, PixelsPerInch = 100 }));
        Assert.Equal("true", (string?)LabelBackground(allSvg, connector.Id).Attribute("data-officeimo-label-adjusted"));
    }

    [Fact]
    public void FluentExportsAndOptionsSnapshotsPreserveLayerSelection() {
        VisioDocument document = CreateLayerDocument();
        VisioPage page = document.Pages[0];
        var options = new VisioImageExportOptions { LayerMode = VisioLayerRenderMode.All, RenderText = false };
        VisioPageImageExportBuilder snapshot = page.ToImage(options);
        options.LayerMode = VisioLayerRenderMode.Visible;
        AssertLayerShapes(ReadSvg(snapshot.AsSvg().Export()), VisioLayerRenderMode.All);
        AssertLayerShapes(ReadSvg(page.ToImage().LayerMode(VisioLayerRenderMode.Printable).AsSvg().Export()), VisioLayerRenderMode.Printable);
        AssertLayerShapes(ReadSvg(document.ToImage().LayerMode(VisioLayerRenderMode.All).AsSvg().Export()), VisioLayerRenderMode.All);
        OfficeImageExportResult batchPage = Assert.Single(document.ToImages().LayerMode(VisioLayerRenderMode.Printable).AsSvg().Export());
        AssertLayerShapes(ReadSvg(batchPage), VisioLayerRenderMode.Printable);
    }

    [Fact]
    public void InvalidLayerSelectionIsRejectedByEveryExportSurface() {
        VisioPage page = VisioDocument.Create().AddPage("Invalid");
        const VisioLayerRenderMode invalid = (VisioLayerRenderMode)99;
        Assert.Throws<ArgumentOutOfRangeException>(() => page.ToSvg(new VisioSvgSaveOptions { LayerMode = invalid }));
        Assert.Throws<ArgumentOutOfRangeException>(() => page.ToPng(new VisioPngSaveOptions { LayerMode = invalid }));
        Assert.Throws<ArgumentOutOfRangeException>(() => page.ExportImage(OfficeImageExportFormat.Svg, new VisioImageExportOptions { LayerMode = invalid }));
        Assert.Throws<ArgumentOutOfRangeException>(() => page.ExportImage(OfficeImageExportFormat.Png, new VisioImageExportOptions { LayerMode = invalid }));
    }

    private static VisioDocument CreateLayerDocument() {
        VisioDocument document = VisioDocument.Create();
        VisioPage page = document.AddPage("Layer flags").Size(6, 2);
        page.AddLayer("Disabled").Visible = false;
        page.FindLayer("Disabled")!.Print = false;
        page.AddLayer("Screen").Print = false;
        page.AddLayer("Print").Visible = false;
        page.AddToLayer("Disabled", AddColoredShape(page, "Disabled", 0.6, OfficeColor.Red));
        page.AddToLayer("Screen", AddColoredShape(page, "Screen", 1.8, OfficeColor.Blue));
        page.AddToLayer("Print", AddColoredShape(page, "Print", 3, OfficeColor.Green));
        VisioShape mixed = AddColoredShape(page, "Mixed", 4.2, OfficeColor.FromRgb(255, 165, 0));
        page.AddToLayer("Screen", mixed).AddToLayer("Print", mixed);
        AddColoredShape(page, "Unassigned", 5.4, OfficeColor.FromRgb(128, 0, 128));
        return document;
    }

    private static VisioShape AddColoredShape(VisioPage page, string name, double pinX, OfficeColor color) {
        VisioShape shape = page.AddRectangle(pinX, 1, 0.8, 0.8, name + " text");
        shape.NameU = name; shape.FillColor = color; shape.LinePattern = 0;
        shape.TextStyle = new VisioTextStyle { Size = 5 };
        return shape;
    }

    private static IEnumerable<VisioDocument> NativeVersions(VisioDocument document) {
        yield return document;
        yield return VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value;
        yield return VisioDocument.Load(new MemoryStream(document.ToBytes()));
    }

    private static void AssertLayerShapes(XDocument svg, VisioLayerRenderMode mode) {
        AssertShape(svg, "Disabled", mode == VisioLayerRenderMode.All);
        AssertShape(svg, "Screen", mode != VisioLayerRenderMode.Printable);
        AssertShape(svg, "Print", mode != VisioLayerRenderMode.Visible);
        AssertShape(svg, "Mixed", true);
        AssertShape(svg, "Unassigned", true);
    }

    private static void AssertShape(XDocument svg, string name, bool expected) =>
        Assert.Equal(expected, svg.Descendants(Svg + "g").Any(group => (string?)group.Attribute("data-visio-nameu") == name));

    private static void AssertLayerPixels(OfficeRasterImage image, double scale, VisioLayerRenderMode mode) {
        AssertPixel(image, scale, 2, 0.6, 1.3, mode == VisioLayerRenderMode.All ? OfficeColor.Red : OfficeColor.White);
        AssertPixel(image, scale, 2, 1.8, 1.3, mode == VisioLayerRenderMode.Printable ? OfficeColor.White : OfficeColor.Blue);
        AssertPixel(image, scale, 2, 3, 1.3, mode == VisioLayerRenderMode.Visible ? OfficeColor.White : OfficeColor.Green);
        AssertPixel(image, scale, 2, 4.2, 1.3, OfficeColor.FromRgb(255, 165, 0));
        AssertPixel(image, scale, 2, 5.4, 1.3, OfficeColor.FromRgb(128, 0, 128));
    }

    private static void AssertPixel(OfficeRasterImage image, double scale, double pageHeight, double x, double y, OfficeColor expected) =>
        Assert.Equal(expected, image.GetPixel((int)Math.Round(x * scale), (int)Math.Round((pageHeight - y) * scale)));

    private static XDocument ReadSvg(OfficeImageExportResult result) => XDocument.Parse(Encoding.UTF8.GetString(result.Bytes));

    private static OfficeRasterImage Decode(byte[] bytes) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, out OfficeRasterImage? image), "Rendered PNG must decode.");
        return image!;
    }

    private static XElement LabelBackground(XDocument svg, string connectorId) => Assert.Single(svg.Descendants(Svg + "g")
        .Single(group => (string?)group.Attribute("data-visio-connector-id") == connectorId)
        .Descendants(Svg + "rect"), rect => (string?)rect.Attribute("data-officeimo-connector-label-background") == "true");

    private static double ReadNumber(XElement element, string name) => double.Parse(element.Attribute(name)!.Value, CultureInfo.InvariantCulture);
}
