using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument.Testing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgGradientTests {
    private static readonly OdfColor Red = OdfColor.Parse("#DC3020"), Blue = OdfColor.Parse("#2040E0");
    private static readonly XNamespace LoExt = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";

    [Theory]
    [InlineData(OdfGradientStyle.Linear)]
    [InlineData(OdfGradientStyle.Axial)]
    [InlineData(OdfGradientStyle.Radial)]
    [InlineData(OdfGradientStyle.Ellipsoid)]
    [InlineData(OdfGradientStyle.Square)]
    [InlineData(OdfGradientStyle.Rectangular)]
    public void NativeDefinitionAndBindingSurvivePackageAndFlatRoundTrips(OdfGradientStyle style) {
        var (document, _, shape) = Scene(new OdfGradientPattern(style, Red, Blue, -450, .3, .4, .8, -.2, 1.2));
        shape.GradientStepCount = 12;
        foreach (var read in RoundTrips(document)) {
            var pattern = read.Styles.FindGradient("Paint")!.Pattern;
            Assert.Equal(style, pattern.Style); Assert.Equal(Red, pattern.StartColor); Assert.Equal(Blue, pattern.EndColor);
            Assert.Equal(-450, pattern.AngleDegrees); Assert.Equal(.3, pattern.Border); Assert.Equal(.4, pattern.StartIntensity);
            Assert.Equal(.8, pattern.EndIntensity); Assert.Equal(-.2, pattern.CenterX); Assert.Equal(1.2, pattern.CenterY);
            Assert.Equal("Paint", read.Pages[0].Shapes[0].FillGradientName); Assert.Equal(12, read.Pages[0].Shapes[0].GradientStepCount);
            Assert.Null(read.Pages[0].Shapes[0].FillColor);
        }
    }

    [Fact]
    public void InheritedFillBindingAndStepCountUseCopyOnWriteWithoutChangingPeers() {
        var (document, page, first) = Scene(new OdfGradientPattern(OdfGradientStyle.Linear, Red, Blue));
        var parent = document.Styles.CreateNamed("Parent", OdfStyleFamily.Graphic);
        parent.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "fill", "gradient");
        parent.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "fill-gradient-name", "Paint");
        parent.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "gradient-step-count", "12");
        first.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", parent.Name);
        var peer = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 1, 1)); peer.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", parent.Name);
        document.MarkPartDirty("content.xml");
        first.GradientStepCount = 0; Assert.Equal(0, first.GradientStepCount); Assert.Equal(12, peer.GradientStepCount);
        first.GradientStepCount = null; Assert.Equal(12, first.GradientStepCount);
        first.FillGradientName = null; Assert.Null(first.FillGradientName); Assert.Equal("Paint", peer.FillGradientName); Assert.Null(first.FillColor);
        first.FillColor = Red; Assert.Equal(Red, first.FillColor); Assert.Null(first.FillGradientName);
        first.FillGradientName = "Paint"; Assert.Null(first.FillColor); Assert.Equal("Paint", peer.FillGradientName);
    }

    [Fact]
    public void SharedDefinitionReplacementPreservesMetadataAndEquivalentProducerStops() {
        var (document, page, _) = Scene(new OdfGradientPattern(OdfGradientStyle.Linear, Red, Blue));
        var peer = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 1, 1)); peer.FillGradientName = "Paint";
        XElement raw = Definition(document); raw.SetAttributeValue(OdfNamespaces.Draw + "display-name", "Friendly paint");
        raw.SetAttributeValue(XNamespace.Get("urn:test") + "metadata", "keep");
        raw.Add(Stop(0, Red), Stop(1, Blue)); raw.Elements().First().SetAttributeValue(XNamespace.Get("urn:test") + "metadata", "endpoint");
        document.MarkPartDirty("styles.xml"); document.Styles.FindGradient("Paint")!.Pattern = new OdfGradientPattern(OdfGradientStyle.Axial, Blue, Red, 90, .2);
        Assert.Equal("Friendly paint", raw.Attribute(OdfNamespaces.Draw + "display-name")!.Value); Assert.Equal("keep", raw.Attribute(XNamespace.Get("urn:test") + "metadata")!.Value);
        Assert.Equal(Blue.ToString(), raw.Elements().First().Attribute(LoExt + "color-value")!.Value);
        Assert.Equal("endpoint", raw.Elements().First().Attribute(XNamespace.Get("urn:test") + "metadata")!.Value);
        Assert.All(RoundTrips(document), read => Assert.Equal(OdfGradientStyle.Axial, read.Styles.FindGradient("Paint")!.Pattern.Style));
        Assert.All(page.ToDrawing().Value.Elements.OfType<OfficeDrawingShape>(), paint => Assert.NotNull(paint.Shape.FillGradient));
    }

    [Theory]
    [InlineData("100grad", 90)]
    [InlineData("1.5707963267948966rad", 90)]
    [InlineData("900", 900)]
    [InlineData("-90deg", -90)]
    public void ReadsStandardAngleUnitsWithoutChangingImportedXml(string angle, double degrees) {
        var (document, _, _) = Scene(new OdfGradientPattern(OdfGradientStyle.Linear, Red, Blue)); Definition(document).SetAttributeValue(OdfNamespaces.Draw + "angle", angle);
        string before = document.GetXml("styles.xml").ToString(); Assert.Equal(degrees, document.Styles.FindGradient("Paint")!.Pattern.AngleDegrees, 10);
        Assert.Equal(before, document.GetXml("styles.xml").ToString());
    }

    [Fact]
    public void OmissionsUseNativeDefaultsAndExplicitAuthoringUsesCenteredGeometry() {
        var (document, _, _) = Scene(new OdfGradientPattern(OdfGradientStyle.Radial, Red, Blue)); XElement raw = Definition(document);
        foreach (string name in new[] { "angle", "border", "start-intensity", "end-intensity", "cx", "cy" }) raw.Attribute(OdfNamespaces.Draw + name)!.Remove();
        var imported = document.Styles.FindGradient("Paint")!.Pattern;
        Assert.Equal(0, imported.AngleDegrees); Assert.Equal(0, imported.Border); Assert.Equal(1, imported.StartIntensity); Assert.Equal(1, imported.EndIntensity);
        Assert.Equal(0, imported.CenterX); Assert.Equal(0, imported.CenterY);
        Assert.Equal(.5, new OdfGradientPattern(OdfGradientStyle.Radial, Red, Blue).CenterX);
    }

    [Fact]
    public void InvalidEditsLeaveDefinitionsAndShapeStylesIntact() {
        var (document, _, shape) = Scene(new OdfGradientPattern(OdfGradientStyle.Linear, Red, Blue));
        string before = Xml(document);
        Assert.Throws<ArgumentOutOfRangeException>(() => new OdfGradientPattern(OdfGradientStyle.Linear, Red, Blue, double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OdfGradientPattern(OdfGradientStyle.Linear, Red, Blue, border: 1.1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OdfGradientPattern(OdfGradientStyle.Linear, Red, Blue, endIntensity: -.1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OdfGradientPattern(OdfGradientStyle.Linear, Red, Blue, centerX: double.MaxValue));
        Assert.Throws<ArgumentException>(() => shape.FillGradientName = "Missing"); Assert.Throws<ArgumentException>(() => shape.FillGradientName = "1Bad");
        Assert.Throws<ArgumentOutOfRangeException>(() => shape.GradientStepCount = 2);
        Assert.Throws<InvalidOperationException>(() => document.Styles.CreateGradient("Paint", new OdfGradientPattern(OdfGradientStyle.Linear, Red, Blue)));
        Assert.Throws<ArgumentException>(() => document.Styles.CreateGradient("Bad name", new OdfGradientPattern(OdfGradientStyle.Linear, Red, Blue)));
        Assert.Equal(before, Xml(document));
    }

    [Theory]
    [InlineData("Missing")]
    [InlineData("DuplicateSvg")]
    [InlineData("Percent")]
    [InlineData("Angle")]
    [InlineData("ExtendedStops")]
    [InlineData("FixedSteps")]
    [InlineData("AxialBorder")]
    [InlineData("OffAreaRadial")]
    [InlineData("Ellipsoid")]
    [InlineData("Square")]
    [InlineData("Rectangular")]
    public void UnsupportedActiveFillRetainsStrokeAndSourceAndEnforcesLossPolicy(string kind) {
        var (document, page, shape) = Scene(new OdfGradientPattern(OdfGradientStyle.Linear, Red, Blue)); XElement raw = Definition(document);
        switch (kind) {
            case "Missing": raw.Remove(); break;
            case "DuplicateSvg": raw.Parent!.Add(new XElement(OdfNamespaces.Svg + "linearGradient", new XAttribute(OdfNamespaces.Draw + "name", "Paint"))); break;
            case "Percent": raw.SetAttributeValue(OdfNamespaces.Draw + "border", "1e1%"); break;
            case "Angle": raw.SetAttributeValue(OdfNamespaces.Draw + "angle", "INFdeg"); break;
            case "ExtendedStops": raw.Add(Stop(0, Red), Stop(.5, Blue), Stop(1, Red)); break;
            case "FixedSteps": shape.GradientStepCount = 12; break;
            case "AxialBorder": document.Styles.FindGradient("Paint")!.Pattern = new OdfGradientPattern(OdfGradientStyle.Axial, Red, Blue, border: 1); break;
            case "OffAreaRadial": document.Styles.FindGradient("Paint")!.Pattern = new OdfGradientPattern(OdfGradientStyle.Radial, Red, Blue, centerX: -.1); break;
            default: raw.SetAttributeValue(OdfNamespaces.Draw + "style", kind.ToLowerInvariant()); break;
        }
        document.MarkPartDirty("styles.xml"); string before = Xml(document); var result = page.ToDrawing();
        OfficeShape paint = Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>()).Shape;
        Assert.NotNull(paint.StrokeColor); Assert.Null(paint.FillGradient); Assert.Null(paint.FillRadialGradient); Assert.Null(paint.FillColor);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "shape:Gradient:fill-gradient" && mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported)); Assert.Equal(before, Xml(document));
    }

    [Fact]
    public void UnsupportedExtendedStopsCannotBeSilentlyOverwrittenOrAttached() {
        var (document, _, shape) = Scene(new OdfGradientPattern(OdfGradientStyle.Linear, Red, Blue)); Definition(document).Add(Stop(.5, Blue));
        document.MarkPartDirty("styles.xml"); string before = Xml(document);
        Assert.Throws<NotSupportedException>(() => document.Styles.FindGradient("Paint")!.Pattern = new OdfGradientPattern(OdfGradientStyle.Linear, Blue, Red));
        Assert.Throws<NotSupportedException>(() => shape.FillGradientName = "Paint"); Assert.Equal(before, Xml(document));
    }

    [Fact]
    public void OpenPathDoesNotResolveAnInactiveBrokenGradient() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var path = page.Shapes.AddPath(new OdfRect(P(10), P(10), P(100), P(60)), new OdfViewBox(0, 0, 100, 60), "M0 0L100 60");
        path.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "fill", "gradient");
        path.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "fill-gradient-name", "Missing");
        var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        Assert.False(result.Report.HasSkippedOrUnsupported); Assert.Contains(result.Report.Mappings, mapping => mapping.Feature.EndsWith(":inactive-fill"));
        Assert.Null(Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>()).Shape.FillGradient);
    }

    [Fact]
    public void LegacyUnitlessAnglesAreReportedRatherThanGuessed() {
        var document = new OdgDocument(OdfPackage.Create(OdfDocumentKind.Graphics, OdfVersion.V1_2), null); var page = document.AddPage();
        document.Styles.CreateGradient("Paint", new OdfGradientPattern(OdfGradientStyle.Linear, Red, Blue)); var shape = page.Shapes.AddRectangle(new OdfRect(P(40), P(40), P(160), P(60)), "Gradient"); shape.FillGradientName = "Paint";
        Definition(document).SetAttributeValue(OdfNamespaces.Draw + "angle", "900"); document.MarkPartDirty("styles.xml");
        Assert.True(page.ToDrawing().Report.HasSkippedOrUnsupported);
        Assert.Throws<InvalidOperationException>(() => document.ToBytes());
        Assert.Throws<NotSupportedException>(() => OdgDocument.Load(new MemoryStream(document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource }))).Styles.FindGradient("Paint")!.Pattern);
        Definition(document).SetAttributeValue(OdfNamespaces.Draw + "angle", "90deg"); Assert.False(page.ToDrawing().Report.HasSkippedOrUnsupported);
        Assert.Equal(90, OdgDocument.Load(new MemoryStream(document.ToBytes())).Styles.FindGradient("Paint")!.Pattern.AngleDegrees);
    }

    [Theory]
    [InlineData(OdfGradientStyle.Linear, 45, 0, 45, 70, 190, 51, 62)]
    [InlineData(OdfGradientStyle.Linear, 30, .3, 160, 70, 125, 56, 128)]
    [InlineData(OdfGradientStyle.Axial, 45, .3, 80, 70, 122, 56, 130)]
    [InlineData(OdfGradientStyle.Radial, 0, 0, 80, 70, 119, 56, 134)]
    public void PaintMatchesNativeSamplesOnANonsquareShape(OdfGradientStyle style, double angle, double border, int x, int y, int r, int g, int b) {
        var (_, page, shape) = Scene(new OdfGradientPattern(style, Red, Blue, angle, border)); shape.StrokeColor = null;
        var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported); var pixel = OfficeDrawingRasterRenderer.Render(result.Value).GetPixel(x, y);
        Assert.InRange(Math.Abs(pixel.R - r), 0, 4); Assert.InRange(Math.Abs(pixel.G - g), 0, 4); Assert.InRange(Math.Abs(pixel.B - b), 0, 4);
        string svg = OfficeDrawingSvgExporter.ToSvg(result.Value); Assert.Contains(style == OdfGradientStyle.Radial ? "radialGradient" : "linearGradient", svg);
    }

    [Fact]
    public void RadialPaintUsesPhysicalCircularRadiusAndKeepsUniformFillOpacity() {
        var (_, page, shape) = Scene(new OdfGradientPattern(OdfGradientStyle.Radial, Red, Blue, border: .3, endIntensity: .4)); shape.StrokeColor = null; shape.FillOpacity = .5;
        var bitmap = OfficeDrawingRasterRenderer.Render(page.ToDrawing().Value); var pixel = bitmap.GetPixel(120, 70);
        Assert.InRange(pixel.A, 126, 129); Assert.InRange(pixel.B, 86, 94);
        var gradient = Assert.Single(page.ToDrawing().Value.Elements.OfType<OfficeDrawingShape>()).Shape.FillRadialGradient!;
        Assert.Equal(gradient.EndRadiusX * 160, gradient.EndRadiusY * 60, 8);
    }

    [Theory]
    [InlineData("rotate(0.25)", true)]
    [InlineData("scale(.7 1.3)", true)]
    [InlineData("translate(10pt 10pt)", false)]
    [InlineData("scale(.7 .7)", false)]
    public void AffineNativeRefittingLimitsAreReportedWithoutDiscardingTheDeclaredGradient(string transform, bool unsupported) {
        var (_, page, shape) = Scene(new OdfGradientPattern(OdfGradientStyle.Linear, Red, Blue, 45)); shape.Transform = transform;
        var result = page.ToDrawing(); Assert.Equal(unsupported, result.Report.HasSkippedOrUnsupported);
        if (unsupported) {
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "shape:Gradient:fill-gradient-transform" && mapping.Status == OdfConversionMappingStatus.Unsupported);
            Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        }
        Assert.Contains("linearGradient", OfficeDrawingSvgExporter.ToSvg(result.Value));
    }

    [Theory]
    [InlineData(1e-30)]
    [InlineData(-1e20)]
    public void CenterPercentagesUseDecimalNotationAcrossTheirFiniteRange(double center) {
        var (document, _, _) = Scene(new OdfGradientPattern(OdfGradientStyle.Linear, Red, Blue, centerX: center));
        string lexical = Definition(document).Attribute(OdfNamespaces.Draw + "cx")!.Value;
        Assert.DoesNotContain("E", lexical); Assert.Equal(center, document.Styles.FindGradient("Paint")!.Pattern.CenterX);
    }

    [Fact]
    public void ReadsIndependentProducerGradientAndPreservesItThroughAnUnrelatedEdit() {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-sheared-geometry.odg"));
        var document = OdgDocument.Load(new MemoryStream(OdfTestPackageRewriter.Rewrite(source)));
        var gradient = document.Styles.FindGradient("Filled")!; Assert.Throws<NotSupportedException>(() => gradient.Pattern);
        Assert.Equal(OdfGradientStyle.Rectangular, document.Styles.FindGradient("Shapes")!.Pattern.Style);
        string before = document.GetXml("styles.xml").ToString(); document.Pages[0].Shapes[0].Name = "Edited name";
        Assert.Equal(before, document.GetXml("styles.xml").ToString()); var read = OdgDocument.Load(new MemoryStream(document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource })));
        Assert.Throws<NotSupportedException>(() => read.Styles.FindGradient("Filled")!.Pattern);
        Assert.Equal(before, read.GetXml("styles.xml").ToString());
    }

    [Theory]
    [InlineData("gradient")]
    [InlineData("opacity")]
    public void VersionUpgradeRejectsAmbiguousNativeAnglesBeforeWritingAndFlatOutputPreservesThem(string kind) {
        var document = new OdgDocument(OdfPackage.Create(OdfDocumentKind.Graphics, OdfVersion.V1_2), null);
        document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(new XElement(OdfNamespaces.Draw + kind,
            new XAttribute(OdfNamespaces.Draw + "name", "Old"), new XAttribute(OdfNamespaces.Draw + "style", "linear"), new XAttribute(OdfNamespaces.Draw + "angle", "900")));
        document.MarkPartDirty("styles.xml"); string before = Xml(document); using var output = new MemoryStream();
        Assert.Throws<InvalidOperationException>(() => document.Save(output)); Assert.Equal(0, output.Length); Assert.Equal(before, Xml(document));
        Assert.Equal("1.2", document.ToFlatXml().Root!.Attribute(OdfNamespaces.Office + "version")!.Value);
        Assert.Equal("900", document.ToFlatXml().Descendants(OdfNamespaces.Draw + kind).Single().Attribute(OdfNamespaces.Draw + "angle")!.Value);
    }

    [Fact]
    public void SharedPresentationModelKeepsGradientDefinitionAndBinding() {
        var presentation = OdpPresentation.Create(); presentation.Styles.CreateGradient("Paint", new OdfGradientPattern(OdfGradientStyle.Radial, Red, Blue));
        presentation.AddSlide("Gradient").AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2)).FillGradientName = "Paint";
        using var flat = new MemoryStream(); presentation.SaveFlatXml(flat); flat.Position = 0;
        var read = OdpPresentation.LoadFlatXml(flat); Assert.Equal("Paint", read.Slides[0].Shapes[0].FillGradientName); Assert.Equal(OdfGradientStyle.Radial, read.Styles.FindGradient("Paint")!.Pattern.Style);
    }

    private static (OdgDocument Document, OdgPage Page, OdgShape Shape) Scene(OdfGradientPattern pattern) {
        var document = OdgDocument.Create(); var page = document.AddPage("Gradient", P(240), P(140)); document.Styles.CreateGradient("Paint", pattern);
        var shape = page.Shapes.AddRectangle(new OdfRect(P(40), P(40), P(160), P(60)), "Gradient"); shape.FillGradientName = "Paint"; return (document, page, shape);
    }
    private static OdfLength P(double value) => OdfLength.Points(value);
    private static XElement Definition(OdgDocument document) => document.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "gradient").Single();
    private static XElement Stop(double offset, OdfColor color) => new XElement(LoExt + "gradient-stop", new XAttribute(OdfNamespaces.Svg + "offset", offset), new XAttribute(LoExt + "color-type", "rgb"), new XAttribute(LoExt + "color-value", color.ToString()));
    private static string Xml(OdgDocument document) => document.GetXml("content.xml").ToString() + document.GetXml("styles.xml");
    private static OdgDocument[] RoundTrips(OdgDocument document) {
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { document, OdgDocument.Load(new MemoryStream(document.ToBytes())), OdgDocument.LoadFlatXml(flat) };
    }
}
