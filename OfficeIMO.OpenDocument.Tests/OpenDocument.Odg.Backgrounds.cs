using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed partial class OpenDocumentOdgBackgroundTests {
    private static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Theory]
    [InlineData(null)]
    [InlineData("none")]
    public void PageNoFillRetainsTheMasterBackgroundBeneathArtwork(string? pageFill) {
        var document = Drawing(); var page = document.Pages[0];
        var master = Bind(document, page.Master!, "styles.xml", "Shared", "solid", "#e07020");
        master.SetAttributeValue(OdfNamespaces.Draw + "background-size", "full");
        if (pageFill != null) Bind(document, page.Element, "content.xml", "Shared", pageFill, "#2040e0");
        page.Shapes.AddRectangle(new OdfRect(P(20), P(20), P(20), P(20))).FillColor = OdfColor.Parse("#00ff00");
        foreach (var read in RoundTrips(document)) {
            string[] before = Parts(read); var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var background = Assert.IsType<OfficeDrawingShape>(result.Value.Elements.First());
            Assert.Equal(OfficeColor.Parse("#e07020"), background.Shape.FillColor);
            var pixels = OfficeDrawingRasterRenderer.Render(result.Value);
            Assert.Equal(OfficeColor.Parse("#e07020"), pixels.GetPixel(0, 0)); Assert.Equal(OfficeColor.Parse("#00ff00"), pixels.GetPixel(30, 30));
            Assert.Equal(before, Parts(read));
        }
    }

    [Fact]
    public void PagePaintResolvesItsOwnStyleScopeButUsesTheMasterBorderArea() {
        var document = Drawing(); var page = document.Pages[0];
        Bind(document, page.Master!, "styles.xml", "Shared", "solid", "#e07020");
        var specific = Bind(document, page.Element, "content.xml", "Shared", "solid", "#2040e0");
        specific.SetAttributeValue(OdfNamespaces.Draw + "background-size", "full"); SetMargins(page, "10pt", "20pt", "15pt", "5pt");
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing(); var background = Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>());
            Assert.Equal(OfficeColor.Parse("#2040e0"), background.Shape.FillColor);
            Assert.Equal(10, background.X); Assert.Equal(15, background.Y); Assert.Equal(70, background.Shape.Width); Assert.Equal(60, background.Shape.Height);
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "page-background:page-area" && mapping.Status == OdfConversionMappingStatus.Unsupported);
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        }
    }

    [Fact]
    public void StretchedEmbeddedBackgroundRetainsBytesAndBorderPlacementAcrossContainers() {
        var document = Drawing(); var page = document.Pages[0];
        var properties = Bitmap(document); SetMargins(page, "10pt", "20pt", "15pt", "5pt"); properties.SetAttributeValue(OdfNamespaces.Draw + "opacity", "50%");
        foreach (var read in RoundTrips(document)) {
            string[] before = Parts(read); var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var background = Assert.Single(result.Value.Elements.OfType<OfficeDrawingImage>());
            Assert.Equal(Png, background.Bytes); Assert.Equal(10, background.Projection.Placement.X); Assert.Equal(15, background.Projection.Placement.Y);
            Assert.Equal(70, background.Projection.Placement.Width); Assert.Equal(60, background.Projection.Placement.Height); Assert.Equal(.5, background.Opacity);
            Assert.False(background.Interpolate);
            Assert.Equal(before, Parts(read));
        }
    }

    [Fact]
    public void IndependentMasterBitmapSurvivesRemovingThePageGradientAndBothContainerSaves() {
        var document = OdgDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-master-artwork.odg"));
        // The producer file remains intact; this edit exposes its existing master bitmap.
        document.Pages[0].Background.UseNoFill();
        var definition = document.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "fill-image")
            .Single(element => (string?)element.Attribute(OdfNamespaces.Draw + "name") == "background");
        byte[] expected = document.GetPackageEntryBytes((string)definition.Attribute(OdfNamespaces.XLink + "href")!);
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing(); var background = Assert.IsType<OfficeDrawingImage>(result.Value.Elements.First());
            Assert.Equal(expected, background.Bytes);
            Assert.Equal(OdfLength.Centimeters(1).ToPoints(), background.Projection.Placement.X, 6);
            Assert.Equal(OdfLength.Centimeters(19).ToPoints(), background.Projection.Placement.Width, 6);
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "page-background" && mapping.Status == OdfConversionMappingStatus.Converted);
            Assert.Contains(result.Report.Mappings, mapping => mapping.Status == OdfConversionMappingStatus.Skipped && mapping.Message!.Contains("custom-shape"));
        }
    }

    [Fact]
    public void StretchModeKeepsInactiveBitmapSizingAndTileDeclarationsWithoutUsingThem() {
        var document = Drawing(); var properties = Bitmap(document);
        properties.SetAttributeValue(OdfNamespaces.Draw + "fill-image-width", "10%");
        properties.SetAttributeValue(OdfNamespaces.Draw + "fill-image-height", "3pt");
        properties.SetAttributeValue(OdfNamespaces.Draw + "fill-image-ref-point", "bottom-right");
        properties.SetAttributeValue(OdfNamespaces.Draw + "tile-repeat-offset", "50% horizontal");
        string before = properties.ToString();
        var background = Assert.Single(document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value.Elements.OfType<OfficeDrawingImage>());
        Assert.Equal(100, background.Projection.Placement.Width); Assert.Equal(80, background.Projection.Placement.Height);
        Assert.Equal(before, properties.ToString());
    }

    [Theory]
    [InlineData("gradient")]
    [InlineData("hatch")]
    public void UnsupportedBackgroundPaintDoesNotRemovePageShapes(string fill) {
        var document = Drawing(); var page = document.Pages[0]; Bind(document, page.Master!, "styles.xml", "Background", fill, "#ff0000");
        page.Shapes.AddRectangle(new OdfRect(P(20), P(20), P(20), P(20)));
        AssertOmittedBackground(document);
    }

    [Theory]
    [InlineData("repeat")]
    [InlineData("no-repeat")]
    public void BitmapRepeatModesAreReportedWithoutStretchFallback(string repeat) {
        var document = Drawing(); Bitmap(document).SetAttributeValue(OdfNamespaces.Style + "repeat", repeat);
        var result = document.Pages[0].ToDrawing(); Assert.Empty(result.Value.Elements);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "page-background" && mapping.Status == OdfConversionMappingStatus.Skipped);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AmbiguousAndExternalFillImagesArePreservedWithOmissionDiagnostics(bool duplicate) {
        var document = Drawing(); Bitmap(document);
        var definition = document.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "fill-image").Single();
        if (duplicate) definition.AddAfterSelf(new XElement(definition));
        else definition.SetAttributeValue(OdfNamespaces.XLink + "href", "https://example.invalid/background.png");
        string[] before = Parts(document); var result = document.Pages[0].ToDrawing(); Assert.Empty(result.Value.Elements);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "page-background" && mapping.Status == OdfConversionMappingStatus.Skipped);
        Assert.Equal(before, Parts(document));
    }

    [Theory]
    [InlineData("-1pt")]
    [InlineData("10%")]
    [InlineData("200pt")]
    public void UnqualifiedBorderMarginsKeepArtworkAndReportTheOmittedBackground(string margin) {
        var document = Drawing(); var page = document.Pages[0]; Bind(document, page.Master!, "styles.xml", "Background", "solid", "#ff0000");
        SetMargins(page, margin, "0pt", "0pt", "0pt"); page.Shapes.AddRectangle(new OdfRect(P(20), P(20), P(20), P(20)));
        AssertOmittedBackground(document);
    }

    private static void AssertOmittedBackground(OdgDocument document) {
        string[] before = Parts(document); var result = document.Pages[0].ToDrawing(); Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "page-background" && mapping.Status == OdfConversionMappingStatus.Skipped);
        Assert.Equal(before, Parts(document));
        Assert.Throws<OdfConversionLossException>(() => document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }
    private static OdgDocument Drawing() { var document = OdgDocument.Create(); document.AddPage("Background", P(100), P(80)); return document; }
    private static XElement Bind(OdgDocument document, XElement owner, string part, string name, string fill, string? color = null) {
        var properties = new XElement(OdfNamespaces.Style + "drawing-page-properties", new XAttribute(OdfNamespaces.Draw + "fill", fill));
        if (color != null) properties.SetAttributeValue(OdfNamespaces.Draw + "fill-color", color);
        document.GetXml(part).Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Add(new XElement(OdfNamespaces.Style + "style",
            new XAttribute(OdfNamespaces.Style + "name", name), new XAttribute(OdfNamespaces.Style + "family", "drawing-page"), properties));
        owner.SetAttributeValue(OdfNamespaces.Draw + "style-name", name); document.MarkPartDirty(part); return properties;
    }
    private static XElement Bitmap(OdgDocument document) {
        string path = OdfImageStore.Add(document, Png, "background.png");
        document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(new XElement(OdfNamespaces.Draw + "fill-image",
            new XAttribute(OdfNamespaces.Draw + "name", "Bitmap"), new XAttribute(OdfNamespaces.XLink + "href", path)));
        var properties = Bind(document, document.Pages[0].Master!, "styles.xml", "Background", "bitmap");
        properties.SetAttributeValue(OdfNamespaces.Draw + "fill-image-name", "Bitmap"); properties.SetAttributeValue(OdfNamespaces.Style + "repeat", "stretch"); return properties;
    }
    private static void SetMargins(OdgPage page, string left, string right, string top, string bottom) {
        XElement properties = page.Element.Document == null ? throw new InvalidOperationException() : page.Master!.Document!.Descendants(OdfNamespaces.Style + "page-layout-properties").Single();
        foreach (var side in new[] { ("left", left), ("right", right), ("top", top), ("bottom", bottom) }) properties.SetAttributeValue(OdfNamespaces.Fo + "margin-" + side.Item1, side.Item2);
    }
    private static OdfLength P(double value) => OdfLength.Points(value);
    private static string[] Parts(OdgDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString() };
    private static OdgDocument[] RoundTrips(OdgDocument document) {
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { document, OdgDocument.Load(new MemoryStream(document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource }))), OdgDocument.LoadFlatXml(flat) };
    }
}
