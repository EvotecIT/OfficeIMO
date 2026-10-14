using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgMasterArtworkTests {
    [Fact]
    public void IndependentMasterTextProjectsWhileFormulaArtworkAndSquareGradientBackgroundRemainReported() {
        var document = OdgDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-master-artwork.odg"));
        string expected = new OdgShapes(document, document.Pages[0].Master!).Last().Text;
        Assert.StartsWith("Sequi aliquid", expected);
        foreach (var read in RoundTrips(document)) {
            string[] before = Parts(read); var result = read.Pages[0].ToDrawing();
            var masterText = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>(), text => text.PlainText == expected);
            Assert.Equal(OdfLength.Centimeters(11.75).ToPoints(), masterText.X, 6);
            Assert.Equal(OdfLength.Centimeters(6.25).ToPoints(), masterText.Y, 6);
            Assert.Empty(result.Value.Elements.OfType<OfficeDrawingImage>());
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "page-background" && mapping.Status == OdfConversionMappingStatus.Skipped);
            Assert.Contains(result.Report.Mappings, mapping => mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Message!.Contains("custom-shape"));
            Assert.Equal(before, Parts(read));
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        }
    }

    [Fact]
    public void MasterPaintUsesItsOwnStylesAndPrecedesPagePaintAcrossBothContainers() {
        var document = OdgDocument.Create(); var page = document.AddPage("Artwork", P(240), P(180));
        Style(document, "styles.xml", "Paint", "graphic", new XElement(OdfNamespaces.Style + "graphic-properties",
            new XAttribute(OdfNamespaces.Draw + "fill", "solid"), new XAttribute(OdfNamespaces.Draw + "fill-color", "#ff0000"),
            new XAttribute(OdfNamespaces.Draw + "stroke", "none")));
        Style(document, "content.xml", "Paint", "graphic", new XElement(OdfNamespaces.Style + "graphic-properties",
            new XAttribute(OdfNamespaces.Draw + "fill", "solid"), new XAttribute(OdfNamespaces.Draw + "fill-color", "#00ff00"),
            new XAttribute(OdfNamespaces.Draw + "stroke", "none")));
        var art = Rect("Master", "Paint", 10, 10, 100, 100); art.SetAttributeValue(OdfNamespaces.Draw + "z-index", 99);
        page.Master!.Add(art); page.Element.Add(Rect("Page", "Paint", 40, 40, 40, 40)); Mark(document);
        foreach (var read in RoundTrips(document)) {
            string[] before = Parts(read);
            var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var bitmap = OfficeDrawingRasterRenderer.Render(result.Value);
            Assert.Equal(OfficeColor.Parse("#ff0000"), bitmap.GetPixel(20, 20));
            Assert.Equal(OfficeColor.Parse("#00ff00"), bitmap.GetPixel(50, 50));
            Assert.Equal(before, Parts(read)); Assert.Single(read.Pages[0].Shapes);
            Assert.Throws<OperationCanceledException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ReportOnly, false, new CancellationToken(true)));
            Assert.Equal(before, Parts(read));
        }
    }

    [Fact]
    public void MasterTextListsImagesAndConnectorReferencesUseTheMasterScope() {
        var document = OdgDocument.Create(); var page = document.AddPage("Master content", P(240), P(180));
        foreach (string part in new[] { "content.xml", "styles.xml" }) {
            bool master = part == "styles.xml";
            Style(document, part, "Paragraph", "paragraph", new XElement(OdfNamespaces.Style + "text-properties",
                new XAttribute(OdfNamespaces.Fo + "font-size", master ? "10pt" : "30pt"), new XAttribute(OdfNamespaces.Fo + "color", master ? "#112233" : "#ff0000")));
            Style(document, part, "Run", "text", new XElement(OdfNamespaces.Style + "text-properties", new XAttribute(OdfNamespaces.Fo + "font-style", master ? "italic" : "normal")));
            document.GetXml(part).Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Add(
                new XElement(OdfNamespaces.Text + "list-style", new XAttribute(OdfNamespaces.Style + "name", "List"),
                    new XElement(OdfNamespaces.Text + "list-level-style-bullet", new XAttribute(OdfNamespaces.Text + "level", "1"),
                        new XAttribute(OdfNamespaces.Text + "bullet-char", master ? "*" : "+"), new XAttribute(OdfNamespaces.Text + "style-name", "Run"))));
        }
        XElement frame = Rect("Text", null, 10, 50, 120, 50);
        frame.Add(new XElement(OdfNamespaces.Text + "p", new XAttribute(OdfNamespaces.Text + "style-name", "Paragraph"),
            new XElement(OdfNamespaces.Text + "span", new XAttribute(OdfNamespaces.Text + "style-name", "Run"), "Master text")),
            new XElement(OdfNamespaces.Text + "list", new XAttribute(OdfNamespaces.Text + "style-name", "List"),
                new XElement(OdfNamespaces.Text + "list-item", new XElement(OdfNamespaces.Text + "p", new XAttribute(OdfNamespaces.Text + "style-name", "Paragraph"), "List body"))));
        var left = Rect("Target", null, 10, 10, 20, 20); left.SetAttributeValue(XNamespace.Xml + "id", "target");
        var line = new XElement(OdfNamespaces.Draw + "connector", new XAttribute(OdfNamespaces.Draw + "name", "Master connector"),
            new XAttribute(OdfNamespaces.Draw + "type", "line"), new XAttribute(OdfNamespaces.Draw + "start-shape", "target"),
            new XAttribute(OdfNamespaces.Draw + "start-glue-point", "1"), new XAttribute(OdfNamespaces.Svg + "x2", "100pt"), new XAttribute(OdfNamespaces.Svg + "y2", "20pt"));
        var bitmap = new OfficeRasterImage(2, 2); bitmap.SetPixel(0, 0, OfficeColor.Parse("#0000ff"));
        bitmap.SetPixel(1, 0, OfficeColor.Parse("#0000ff")); bitmap.SetPixel(0, 1, OfficeColor.Parse("#0000ff")); bitmap.SetPixel(1, 1, OfficeColor.Parse("#0000ff"));
        OdgShape image = page.Shapes.AddImage(OfficePngWriter.Encode(bitmap), "master.png", new OdfRect(P(150), P(50), P(20), P(20)));
        XElement imageXml = new XElement(image.Element); page.Shapes.RemoveAt(0);
        page.Master!.Add(new XElement(OdfNamespaces.Draw + "g", new XAttribute(OdfNamespaces.Draw + "transform", "translate(10pt 5pt)"), left, line), frame, imageXml);
        XElement unrelated = Rect("Page target", null, 150, 100, 20, 20); unrelated.SetAttributeValue(XNamespace.Xml + "id", "target"); page.Element.Add(unrelated);
        Mark(document);
        foreach (var read in RoundTrips(document)) {
            string[] before = Parts(read); var result = read.Pages[0].ToDrawing();
            Assert.False(result.Report.HasSkippedOrUnsupported);
            OfficeDrawingRichText text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
            OfficeRichTextRun run = text.Paragraphs[0].Runs.Single();
            Assert.Equal("Master text", run.Text); Assert.Equal(10, run.FontSize); Assert.True(run.Italic);
            Assert.Equal(OfficeColor.Parse("#112233"), run.Color);
            Assert.Equal("*", text.Paragraphs[1].Label!.Run.Text); Assert.True(text.Paragraphs[1].Label!.Run.Italic);
            Assert.Single(result.Value.Elements.OfType<OfficeDrawingImage>());
            var connector = new OdgShapes(read, read.Pages[0].Master!)[0].Children[1];
            Assert.Equal(30, connector.X1.ToPoints(), 6); Assert.Equal(20, connector.Y1.ToPoints(), 6);
            var rendered = OfficeDrawingRasterRenderer.Render(result.Value);
            Assert.Equal(OfficeColor.Parse("#0000ff"), rendered.GetPixel(155, 55));
            Assert.Equal(before, Parts(read));
        }
    }

    [Fact]
    public void MasterArtworkUsesEffectiveLayersForScreenAndPrint() {
        var document = OdgDocument.Create(); var source = document.AddPage("Shared"); var other = document.AddPage("Override");
        other.MasterPageName = source.MasterPageName;
        source.MasterLayers.Add("Screen", OdgLayerDisplay.Screen); source.MasterLayers.Add("Print", OdgLayerDisplay.Printer);
        foreach (string layer in new[] { "Screen", "Print" }) {
            XElement shape = Rect(layer, null, 10, layer == "Screen" ? 10 : 50, 100, 30);
            shape.SetAttributeValue(OdfNamespaces.Draw + "layer", layer); shape.Add(new XElement(OdfNamespaces.Text + "p", layer)); source.Master!.Add(shape);
        }
        other.Layers.Add("Screen", OdgLayerDisplay.None); other.Layers.Add("Print", OdgLayerDisplay.None); Mark(document);
        foreach (var read in RoundTrips(document)) {
            Assert.Equal("Screen", Assert.Single(read.Pages[0].ToDrawing().Value.Elements.OfType<OfficeDrawingRichText>()).PlainText);
            Assert.Equal("Print", Assert.Single(read.Pages[0].ToDrawing(forPrint: true).Value.Elements.OfType<OfficeDrawingRichText>()).PlainText);
            Assert.Empty(read.Pages[1].ToDrawing().Value.Elements); Assert.Empty(read.Pages[1].ToDrawing(forPrint: true).Value.Elements);
        }
    }

    [Theory]
    [InlineData("false")]
    [InlineData("0")]
    [InlineData("true")]
    [InlineData("1")]
    [InlineData("invalid")]
    public void PresentationVisibilityPropertiesRemainReportedWithoutHidingDrawArtwork(string value) {
        var document = OdgDocument.Create(); var page = document.AddPage(); page.Master!.Add(Rect("Art", null, 0, 0, 20, 20));
        var parent = document.Styles.CreateNamed("PageParent", OdfStyleFamily.DrawingPage);
        parent.SetProperty(OdfNamespaces.Style + "drawing-page-properties", OdfNamespaces.Presentation + "background-objects-visible", value);
        var child = document.Styles.CreateAutomatic(OdfStyleFamily.DrawingPage, parentStyleName: parent.Name);
        page.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", child.Name); Mark(document);
        foreach (var read in RoundTrips(document)) {
            string[] before = Parts(read); var result = read.Pages[0].ToDrawing(); Assert.Single(result.Value.Elements);
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "page-style" && mapping.Status == OdfConversionMappingStatus.Unsupported);
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            Assert.Equal(before, Parts(read));
        }
    }

    [Fact]
    public void UnsupportedMasterObjectsRemainReportedAndPageShapesStillProject() {
        var document = OdgDocument.Create(); var page = document.AddPage(); page.Master!.Add(new XElement(OdfNamespaces.Draw + "control", new XAttribute(OdfNamespaces.Draw + "name", "Control")));
        page.Shapes.AddRectangle(new OdfRect(P(10), P(10), P(20), P(20))); Mark(document);
        string[] before = Parts(document); var result = page.ToDrawing(); Assert.Single(result.Value.Elements);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "shape:Control" && mapping.Status == OdfConversionMappingStatus.Skipped);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported)); Assert.Equal(before, Parts(document));
    }

    [Fact]
    public void MasterArtworkSurvivesImportAndIndependentMasterCloning() {
        var document = OdgDocument.Create(); var page = document.AddPage("Source");
        Style(document, "styles.xml", "MasterPaint", "graphic", new XElement(OdfNamespaces.Style + "graphic-properties",
            new XAttribute(OdfNamespaces.Draw + "fill", "solid"), new XAttribute(OdfNamespaces.Draw + "fill-color", "#ff0000")));
        page.Master!.Add(Rect("Artwork", "MasterPaint", 20, 20, 40, 40)); Mark(document);
        string expected = OfficeDrawingSvgExporter.ToSvg(page.ToDrawing().Value);
        var cloned = document.ClonePage(0); cloned.MasterPageName = document.CloneMasterPage(page.MasterPageName);
        Assert.Equal(expected, OfficeDrawingSvgExporter.ToSvg(cloned.ToDrawing().Value));
        var destination = OdgDocument.Create(); var imported = destination.ImportPage(document, 0);
        Assert.Equal(expected, OfficeDrawingSvgExporter.ToSvg(imported.ToDrawing().Value));
        foreach (var read in RoundTrips(destination)) Assert.Equal(expected, OfficeDrawingSvgExporter.ToSvg(read.Pages[0].ToDrawing().Value));
    }

    private static OdfLength P(double value) => OdfLength.Points(value);
    private static XElement Rect(string name, string? style, double x, double y, double width, double height) {
        var element = new XElement(OdfNamespaces.Draw + "rect", new XAttribute(OdfNamespaces.Draw + "name", name));
        if (style != null) element.SetAttributeValue(OdfNamespaces.Draw + "style-name", style);
        OdfShape.ApplyBounds(element, new OdfRect(P(x), P(y), P(width), P(height))); return element;
    }
    private static void Style(OdgDocument document, string part, string name, string family, XElement properties) =>
        document.GetXml(part).Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Add(new XElement(OdfNamespaces.Style + "style",
            new XAttribute(OdfNamespaces.Style + "name", name), new XAttribute(OdfNamespaces.Style + "family", family), properties));
    private static string[] Parts(OdgDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString() };
    private static void Mark(OdgDocument document) { document.MarkPartDirty("content.xml"); document.MarkPartDirty("styles.xml"); }
    private static OdgDocument[] RoundTrips(OdgDocument document) {
        byte[] package = document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource });
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { document, OdgDocument.Load(new MemoryStream(package)), OdgDocument.LoadFlatXml(flat) };
    }
}
