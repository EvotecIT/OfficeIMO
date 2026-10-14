using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument.Testing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgStrokeTests {
    [Fact]
    public void AuthorsTwoSequencePercentageDashesWithExplicitCapsAndJoins() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        document.Styles.CreateStrokeDash("Mixed", new OdfStrokeDashPattern(2, OdfLength.Parse("200%"), OdfLength.Parse("100%"), 1, OdfLength.Points(12), true));
        var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2));
        shape.StrokeWidth = OdfLength.Points(4); shape.StrokeDashName = "Mixed";
        shape.StrokeLineCap = OfficeStrokeLineCap.Square; shape.StrokeLineJoin = OfficeStrokeLineJoin.Bevel;
        shape.StrokeColor = OdfColor.Parse("#2244BB");
        foreach (var read in RoundTrips(document)) {
            var actual = read.Pages[0].Shapes[0]; Assert.Equal("Mixed", actual.StrokeDashName);
            OfficeShape painted = Assert.Single(read.Pages[0].ToDrawing().Value.Elements.OfType<OfficeDrawingShape>()).Shape;
            Assert.Equal(new double[] { 8, 4, 8, 4, 12, 4 }, painted.StrokeDashArray);
            Assert.Equal(OfficeStrokeLineCap.Square, painted.StrokeLineCap); Assert.Equal(OfficeStrokeLineJoin.Bevel, painted.StrokeLineJoin);
            Assert.Equal(OfficeColor.Parse("#2244BB"), painted.StrokeColor);
            string svg = OfficeDrawingSvgExporter.ToSvg(read.Pages[0].ToDrawing().Value);
            Assert.Contains("stroke-dasharray=\"8 4 8 4 12 4\"", svg); Assert.Contains("stroke-linecap=\"square\"", svg);
        }
    }

    [Fact]
    public void InheritedDashCapsAndJoinsAndSharedStyleEditsRemainIndependent() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        document.Styles.CreateStrokeDash("Round", new OdfStrokeDashPattern(1, OdfLength.Points(8), OdfLength.Points(6), roundCaps: true));
        var named = document.Styles.CreateNamed("Parent", OdfStyleFamily.Graphic);
        named.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "stroke", "dash");
        named.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "stroke-dash", "Round");
        named.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Svg + "stroke-color", "#000000");
        named.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Svg + "stroke-linecap", "butt");
        named.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "stroke-linejoin", "miter");
        var first = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2));
        first.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", named.Name); document.MarkPartDirty("content.xml");
        first.StrokeLineCap = OfficeStrokeLineCap.Square; first.StrokeLineCap = null;
        Assert.Equal(OfficeStrokeLineCap.Butt, first.StrokeLineCap); Assert.Equal(OfficeStrokeLineJoin.Miter, first.StrokeLineJoin);
        var other = page.Shapes.AddRectangle(OdfRect.FromCentimeters(4, 1, 2, 2));
        other.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", first.Element.Attribute(OdfNamespaces.Draw + "style-name")!.Value); document.MarkPartDirty("content.xml");
        first.StrokeDashName = null; first.StrokeLineJoin = OfficeStrokeLineJoin.Bevel;
        Assert.Null(first.StrokeDashName); Assert.Equal("Round", other.StrokeDashName); Assert.Equal(OfficeStrokeLineJoin.Miter, other.StrokeLineJoin);
        var projected = page.ToDrawing().Value.Elements.OfType<OfficeDrawingShape>().ToArray();
        Assert.Empty(projected[0].Shape.StrokeDashArray); Assert.NotEmpty(projected[1].Shape.StrokeDashArray);
        Assert.Equal(OfficeStrokeLineCap.Butt, projected[1].Shape.StrokeLineCap);
    }

    [Fact]
    public void DefinitionReplacementIsSharedAndPreservesUnknownMetadata() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var dash = document.Styles.CreateStrokeDash("Shared", new OdfStrokeDashPattern(1, OdfLength.Points(8), OdfLength.Points(6), roundCaps: true));
        XElement raw = document.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "stroke-dash").Single();
        raw.SetAttributeValue(OdfNamespaces.Draw + "display-name", "Readable dash"); raw.SetAttributeValue(XNamespace.Get("urn:test") + "metadata", "keep"); document.MarkPartDirty("styles.xml");
        foreach (int x in new[] { 1, 4 }) page.Shapes.AddRectangle(OdfRect.FromCentimeters(x, 1, 2, 2)).StrokeDashName = dash.Name;
        dash.Pattern = new OdfStrokeDashPattern(1, OdfLength.Points(12), OdfLength.Points(4));
        Assert.Equal("Readable dash", raw.Attribute(OdfNamespaces.Draw + "display-name")!.Value); Assert.Equal("keep", raw.Attribute(XNamespace.Get("urn:test") + "metadata")!.Value);
        Assert.All(page.ToDrawing().Value.Elements.OfType<OfficeDrawingShape>(), shape => {
            Assert.Equal(new double[] { 12, 4 }, shape.Shape.StrokeDashArray); Assert.Equal(OfficeStrokeLineCap.Butt, shape.Shape.StrokeLineCap);
        });
    }

    [Theory]
    [InlineData(-1, "8pt", "6pt")]
    [InlineData(1025, "8pt", "6pt")]
    [InlineData(0, "8pt", "6pt")]
    [InlineData(1, "-1pt", "6pt")]
    [InlineData(1, "8pt", "0pt")]
    [InlineData(1, "8pt", "1e2%")]
    [InlineData(1, "+8pt", "6pt")]
    [InlineData(1, "8PT", "6pt")]
    [InlineData(1, "8pt", "+6pt")]
    public void InvalidPatternOrReferenceEditsDoNotMutateDocument(int count, string length, string gap) {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2));
        string before = document.GetXml("content.xml").ToString() + document.GetXml("styles.xml");
        Assert.ThrowsAny<ArgumentException>(() => document.Styles.CreateStrokeDash("Bad", new OdfStrokeDashPattern(count, OdfLength.Parse(length), OdfLength.Parse(gap))));
        Assert.Throws<ArgumentException>(() => shape.StrokeDashName = "Missing");
        Assert.Throws<ArgumentOutOfRangeException>(() => shape.StrokeLineCap = (OfficeStrokeLineCap)99);
        Assert.Throws<ArgumentOutOfRangeException>(() => shape.StrokeLineJoin = (OfficeStrokeLineJoin)99);
        Assert.Equal(before, document.GetXml("content.xml").ToString() + document.GetXml("styles.xml"));
    }

    [Theory]
    [InlineData("Missing")]
    [InlineData("Duplicate")]
    [InlineData("Omitted")]
    [InlineData("Huge")]
    public void InvalidImportedDefinitionsArePreservedAndReported(string kind) {
        var document = OdgDocument.Create(); var page = document.AddPage(); var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2), "Invalid");
        document.Styles.CreateStrokeDash("Pattern", new OdfStrokeDashPattern(1, OdfLength.Points(8), OdfLength.Points(6))); shape.StrokeDashName = "Pattern";
        XElement raw = document.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "stroke-dash").Single();
        switch (kind) {
            case "Missing": raw.Remove(); break;
            case "Duplicate": raw.Parent!.Add(new XElement(raw)); break;
            case "Omitted": raw.Attribute(OdfNamespaces.Draw + "dots1-length")!.Remove(); break;
            case "Huge": raw.SetAttributeValue(OdfNamespaces.Draw + "dots1", int.MaxValue); break;
        }
        document.MarkPartDirty("styles.xml"); string before = document.GetXml("styles.xml").ToString();
        Assert.Empty(page.ToDrawing().Value.Elements); Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal(before, document.GetXml("styles.xml").ToString());
    }

    [Fact]
    public void ExtraDashOverlaysAreReportedWithoutDroppingPrimaryGeometry() {
        var document = OdgDocument.Create(); var page = document.AddPage(); var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2));
        shape.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "stroke-dash-names", "Overlay");
        Assert.Single(page.ToDrawing().Value.Elements); Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }

    [Fact]
    public void SharedPresentationOwnerRetainsDashDefinitionsAndStrokeProperties() {
        var document = OdpPresentation.Create(); document.Styles.CreateStrokeDash("SlideDash", new OdfStrokeDashPattern(1, OdfLength.Points(8), OdfLength.Points(6)));
        var shape = document.AddSlide("Dash").AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2)); shape.StrokeDashName = "SlideDash"; shape.StrokeLineCap = OfficeStrokeLineCap.Round; shape.StrokeLineJoin = OfficeStrokeLineJoin.Miter;
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        foreach (var read in new[] { OdpPresentation.Load(new MemoryStream(document.ToBytes())), OdpPresentation.LoadFlatXml(flat) }) {
            Assert.Equal("SlideDash", read.Slides[0].Shapes[0].StrokeDashName); Assert.Equal(OfficeStrokeLineCap.Round, read.Slides[0].Shapes[0].StrokeLineCap); Assert.Equal(1, read.Styles.FindStrokeDash("SlideDash")!.Pattern.DashCount);
        }
    }
    [Fact]
    public void ReadsIndependentProducerDashDefinitionAndPreservesItOnShapeEdit() {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-network-connectors.odg"));
        var document = OdgDocument.Load(new MemoryStream(OdfTestPackageRewriter.Rewrite(source)));
        var dash = Assert.Single(document.Styles.StrokeDashes); Assert.Equal(1, dash.Pattern.DashCount); Assert.Equal(1, dash.Pattern.SecondDashCount);
        Assert.Equal("0.02cm", dash.Pattern.Distance.ToString()); string stylesBefore = document.GetXml("styles.xml").ToString();
        var line = document.Pages[0].Shapes.AddLine(OdfLength.Points(10), OdfLength.Points(10), OdfLength.Points(100), OdfLength.Points(10)); line.StrokeDashName = dash.Name;
        Assert.Equal(stylesBefore, document.GetXml("styles.xml").ToString());
        var read = OdgDocument.Load(new MemoryStream(document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource }))); Assert.Equal(dash.Name, read.Pages[0].Shapes.Last().StrokeDashName);
    }

    [Theory]
    [InlineData("none")]
    [InlineData("middle")]
    public void PreservesUnsupportedNativeJoinAndReportsIt(string token) {
        var document = OdgDocument.Create(); var page = document.AddPage(); var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2));
        shape.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "stroke-linejoin", token);
        string before = document.GetXml("content.xml").ToString(); Assert.Empty(page.ToDrawing().Value.Elements);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported)); Assert.Equal(before, document.GetXml("content.xml").ToString());
    }

    [Theory]
    [InlineData("Long Short")]
    [InlineData("prefix:name")]
    [InlineData("1Invalid")]
    public void RejectsInvalidNamesAndReferencesAcrossStyleCreationBeforeMutation(string name) {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2));
        // Imported definitions remain inspectable, but cannot be attached using a new invalid reference.
        document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(new XElement(OdfNamespaces.Draw + "stroke-dash", new XAttribute(OdfNamespaces.Draw + "name", name)));
        document.MarkPartDirty("styles.xml"); var named = document.Styles.CreateNamed("ValidParentEdit", OdfStyleFamily.Graphic);
        string before = document.GetXml("content.xml").ToString() + document.GetXml("styles.xml");
        var pattern = new OdfStrokeDashPattern(1, OdfLength.Points(8), OdfLength.Points(4));
        Assert.NotNull(document.Styles.FindStrokeDash(name)); Assert.Throws<ArgumentException>(() => document.Styles.CreateStrokeDash(name, pattern));
        Assert.Throws<ArgumentException>(() => shape.StrokeDashName = name);
        Assert.Throws<ArgumentException>(() => document.Styles.CreateNamed(name, OdfStyleFamily.Graphic));
        Assert.Throws<ArgumentException>(() => named.ParentStyleName = name);
        Assert.Throws<ArgumentException>(() => document.Styles.CreateNamed("Valid", OdfStyleFamily.Graphic, name));
        Assert.Throws<ArgumentException>(() => document.Styles.CreateAutomatic(OdfStyleFamily.Graphic, parentStyleName: name));
        Assert.Equal(before, document.GetXml("content.xml").ToString() + document.GetXml("styles.xml"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DisabledOdgLineStaysInvisibleInSvgAndRaster(bool transformed) {
        var document = OdgDocument.Create(); var page = document.AddPage("Invisible", OdfLength.Points(100), OdfLength.Points(100));
        var line = page.Shapes.AddLine(OdfLength.Points(20), OdfLength.Points(20), OdfLength.Points(70), OdfLength.Points(20)); line.StrokeColor = null;
        if (transformed) line.Transform = "translate(0pt 10pt)";
        var scene = page.ToDrawing().Value; Assert.Contains("stroke=\"none\"", OfficeDrawingSvgExporter.ToSvg(scene));
        var image = OfficeDrawingRasterRenderer.Render(scene); for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++) Assert.Equal(0, image.GetPixel(x, y).A);
    }

    private static OdgDocument[] RoundTrips(OdgDocument document) {
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { document, OdgDocument.Load(new MemoryStream(document.ToBytes())), OdgDocument.LoadFlatXml(flat) };
    }
}
