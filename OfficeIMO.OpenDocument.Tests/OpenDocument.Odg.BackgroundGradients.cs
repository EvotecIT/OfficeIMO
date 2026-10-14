using System;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed partial class OpenDocumentOdgBackgroundTests {
    [Theory]
    [InlineData(OdfGradientStyle.Linear)]
    [InlineData(OdfGradientStyle.Axial)]
    [InlineData(OdfGradientStyle.Radial)]
    public void GradientBackgroundPreservesItsBindingOpacityAndBorderAreaAcrossContainers(OdfGradientStyle style) {
        var document = Drawing(); var page = document.Pages[0];
        document.Styles.CreateGradient("Paint", new OdfGradientPattern(style, OdfColor.Parse("#dc3020"), OdfColor.Parse("#2040e0"), 40, .2, .6, .8, .7, .6));
        var properties = Bind(document, page.Master!, "styles.xml", "Background", "gradient");
        properties.SetAttributeValue(OdfNamespaces.Draw + "fill-gradient-name", "Paint");
        properties.SetAttributeValue(OdfNamespaces.Draw + "opacity", "50%"); SetMargins(page, "10pt", "20pt", "15pt", "5pt");
        page.Shapes.AddRectangle(new OdfRect(P(20), P(20), P(20), P(20))).FillColor = OdfColor.Parse("#00ff00");
        foreach (var read in RoundTrips(document)) {
            string[] before = Parts(read); var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var background = Assert.IsType<OfficeDrawingShape>(result.Value.Elements.First());
            Assert.Equal(10, background.X); Assert.Equal(15, background.Y); Assert.Equal(70, background.Shape.Width); Assert.Equal(60, background.Shape.Height);
            Assert.Equal(.5, background.Shape.FillOpacity); Assert.Null(background.Shape.StrokeColor); Assert.Null(background.Shape.FillColor);
            if (style == OdfGradientStyle.Radial) Assert.NotNull(background.Shape.FillRadialGradient); else Assert.NotNull(background.Shape.FillGradient);
            Assert.Contains(result.Report.Mappings, m => m.Feature == "page-background" && m.Status == OdfConversionMappingStatus.Approximated);
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnAnyLoss));
            Assert.Equal(OfficeColor.Parse("#00ff00"), OfficeDrawingRasterRenderer.Render(result.Value).GetPixel(30, 30));
            Assert.Equal(style, read.Styles.FindGradient("Paint")!.Pattern.Style); Assert.Equal(before, Parts(read));
        }
    }

    [Fact]
    public void PageGradientOverridesMasterColorWithoutResolvingItsInactiveGradientReference() {
        var document = Drawing(); var page = document.Pages[0];
        document.Styles.CreateGradient("Paint", new OdfGradientPattern(OdfGradientStyle.Linear, OdfColor.Parse("#dc3020"), OdfColor.Parse("#2040e0"), 90));
        var master = Bind(document, page.Master!, "styles.xml", "Shared", "solid", "#00ff00");
        master.SetAttributeValue(OdfNamespaces.Draw + "fill-gradient-name", "Missing"); master.SetAttributeValue(OdfNamespaces.Draw + "background-size", "full");
        var selected = Bind(document, page.Element, "content.xml", "Shared", "gradient"); selected.SetAttributeValue(OdfNamespaces.Draw + "fill-gradient-name", "Paint");
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported); var paint = Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>());
            Assert.Equal(0, paint.X); Assert.Equal(0, paint.Y); Assert.Equal(100, paint.Shape.Width); Assert.Equal(80, paint.Shape.Height);
            Assert.NotNull(paint.Shape.FillGradient); Assert.Null(paint.Shape.FillColor);
            Assert.DoesNotContain(result.Report.Mappings, m => m.Status is OdfConversionMappingStatus.Skipped or OdfConversionMappingStatus.Unsupported);
        }
    }

    [Fact]
    public void OneHundredPercentLinearBorderUsesTheIntensityAdjustedStartColor() {
        var document = Drawing(); var page = document.Pages[0];
        document.Styles.CreateGradient("Paint", new OdfGradientPattern(OdfGradientStyle.Linear, OdfColor.Parse("#dc3020"), OdfColor.Parse("#2040e0"), border: 1, startIntensity: .5));
        var properties = Bind(document, page.Master!, "styles.xml", "Background", "gradient"); properties.SetAttributeValue(OdfNamespaces.Draw + "fill-gradient-name", "Paint");
        var paint = Assert.Single(page.ToDrawing().Value.Elements.OfType<OfficeDrawingShape>());
        Assert.Equal(OfficeColor.Parse("#6e1810"), paint.Shape.FillColor); Assert.Null(paint.Shape.FillGradient);
    }

    [Theory]
    [InlineData("MissingBinding")]
    [InlineData("MissingDefinition")]
    [InlineData("DuplicateSvg")]
    [InlineData("FixedBands")]
    [InlineData("InvalidBands")]
    [InlineData("NegativeBands")]
    [InlineData("Square")]
    [InlineData("ExtendedStops")]
    [InlineData("OffAreaRadial")]
    [InlineData("SingularAxial")]
    [InlineData("OpacityGradient")]
    public void UnsupportedBackgroundGradientsPreserveTheirSourceAndKeepPageArtwork(string kind) {
        var document = Drawing(); var page = document.Pages[0];
        document.Styles.CreateGradient("Paint", new OdfGradientPattern(OdfGradientStyle.Linear, OdfColor.Parse("#dc3020"), OdfColor.Parse("#2040e0")));
        var properties = Bind(document, page.Master!, "styles.xml", "Background", "gradient"); properties.SetAttributeValue(OdfNamespaces.Draw + "fill-gradient-name", "Paint");
        var raw = document.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "gradient").Single();
        switch (kind) {
            case "MissingBinding": properties.Attribute(OdfNamespaces.Draw + "fill-gradient-name")!.Remove(); break;
            case "MissingDefinition": raw.Remove(); break;
            case "DuplicateSvg": raw.AddAfterSelf(new XElement(OdfNamespaces.Svg + "linearGradient", new XAttribute(OdfNamespaces.Draw + "name", "Paint"))); break;
            case "FixedBands": properties.SetAttributeValue(OdfNamespaces.Draw + "gradient-step-count", "12"); break;
            case "InvalidBands": properties.SetAttributeValue(OdfNamespaces.Draw + "gradient-step-count", "2"); break;
            case "NegativeBands": properties.SetAttributeValue(OdfNamespaces.Draw + "gradient-step-count", "-1"); break;
            case "Square": raw.SetAttributeValue(OdfNamespaces.Draw + "style", "square"); break;
            case "ExtendedStops": raw.Add(new XElement(XNamespace.Get("urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0") + "gradient-stop")); break;
            case "OffAreaRadial": document.Styles.FindGradient("Paint")!.Pattern = new OdfGradientPattern(OdfGradientStyle.Radial, OdfColor.Parse("#dc3020"), OdfColor.Parse("#2040e0"), centerX: -.1); break;
            case "SingularAxial": document.Styles.FindGradient("Paint")!.Pattern = new OdfGradientPattern(OdfGradientStyle.Axial, OdfColor.Parse("#dc3020"), OdfColor.Parse("#2040e0"), border: 1); break;
            case "OpacityGradient": properties.SetAttributeValue(OdfNamespaces.Draw + "opacity-name", "Opacity"); break;
        }
        document.MarkPartDirty("styles.xml"); page.Shapes.AddRectangle(new OdfRect(P(20), P(20), P(20), P(20)));
        foreach (var read in RoundTrips(document)) AssertOmittedBackground(read);
    }
}
