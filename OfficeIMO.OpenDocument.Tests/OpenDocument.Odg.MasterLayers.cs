using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgMasterLayerTests {
    [Fact]
    public void SharesMasterLayersAndLayoutWhilePageOverridesAndOtherMastersRemainIndependent() {
        var document = OdgDocument.Create();
        document.Layers.Add("Review", OdgLayerDisplay.None);
        document.Layers.Add("Output", OdgLayerDisplay.None);
        var source = document.AddPage("Source", OdfLength.Points(300), OdfLength.Points(200));
        var shared = document.AddPage("Shared");
        var overridden = document.AddPage("Override");
        var other = document.AddPage("Other");
        source.MasterLayers.Add("Review", OdgLayerDisplay.Screen).IsProtected = true;
        source.MasterLayers.Add("Output", OdgLayerDisplay.Printer);
        shared.MasterPageName = source.MasterPageName;
        overridden.MasterPageName = source.MasterPageName;
        overridden.Layers.Add("Review", OdgLayerDisplay.Printer);
        overridden.Layers.Add("Output", OdgLayerDisplay.Screen);
        foreach (OdgPage page in document.Pages) {
            foreach (string name in new[] { "Review", "Output" }) {
                var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 2, 2), name);
                shape.Text = name; shape.Layer = name;
            }
        }
        shared.Width = OdfLength.Points(400);
        foreach (var read in RoundTrips(document)) {
            Assert.Equal(400, read.Pages[0].Width.ToPoints(), 3);
            Assert.Equal(400, read.Pages[1].Width.ToPoints(), 3);
            Assert.Equal(400, read.Pages[2].Width.ToPoints(), 3);
            Assert.Equal(200, read.Pages[1].Height.ToPoints(), 3);
            Assert.NotEqual(400, read.Pages[3].Width.ToPoints());
            Assert.True(read.Pages[1].EffectiveLayers.Find("Review")!.IsProtected);
            Assert.False(read.Pages[2].EffectiveLayers.Find("Review")!.IsProtected);
            Assert.Equal(OdgLayerDisplay.None, read.Layers.Find("Review")!.Display);
            Assert.Empty(read.Pages[3].MasterLayers);
            AssertLabels(read.Pages[0], "Review", "Output");
            AssertLabels(read.Pages[1], "Review", "Output");
            AssertLabels(read.Pages[2], "Output", "Review");
            AssertLabels(read.Pages[3], null, null);
            read.Pages[1].MasterLayers.Find("Review")!.Display = OdgLayerDisplay.None;
            AssertLabels(read.Pages[0], null, "Output");
            AssertLabels(read.Pages[2], "Output", "Review");
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("missing")]
    public void RejectsUnknownMasterWithoutChangingThePageOrStyles(string name) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        string[] before = Parts(document);
        Assert.Throws<ArgumentException>(() => page.MasterPageName = name);
        Assert.Equal(before, Parts(document));
    }

    [Theory]
    [InlineData("duplicate-master")]
    [InlineData("missing-layout")]
    [InlineData("duplicate-layout")]
    public void RejectsUnresolvableMasterLayoutBeforeRebindingThePage(string defect) {
        var document = OdgDocument.Create(); var page = document.AddPage(); var target = document.AddPage();
        XElement master = target.Master!;
        if (defect == "duplicate-master") master.AddAfterSelf(new XElement(master));
        else if (defect == "missing-layout") master.SetAttributeValue(OdfNamespaces.Style + "page-layout-name", "missing");
        else {
            XElement layout = document.GetXml("styles.xml").Descendants(OdfNamespaces.Style + "page-layout")
                .Single(element => (string?)element.Attribute(OdfNamespaces.Style + "name") == (string?)master.Attribute(OdfNamespaces.Style + "page-layout-name"));
            layout.AddAfterSelf(new XElement(layout));
        }
        string[] before = Parts(document);
        Assert.Throws<InvalidDataException>(() => page.MasterPageName = target.MasterPageName);
        Assert.Equal(before, Parts(document));
    }

    [Fact]
    public void MasterLayerAccessRejectsDanglingReferencesWithoutDeclaringADifferentScope() {
        var document = OdgDocument.Create(); var page = document.AddPage();
        page.Element.SetAttributeValue(OdfNamespaces.Draw + "master-page-name", "missing");
        string[] before = Parts(document);
        Assert.Throws<InvalidDataException>(() => page.MasterLayers.Add("Review"));
        Assert.Equal(before, Parts(document));
        Assert.Empty(document.Layers); Assert.Empty(page.Layers);
    }

    [Fact]
    public void EditsIndependentProducerMasterWithoutChangingGlobalSavedViewsOrMasterArtwork() {
        var document = OdgDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-layers.odg"));
        var page = document.Pages[0];
        var settings = new XElement(document.GetXml("settings.xml").Root!.Element(OdfNamespaces.Office + "settings")!);
        string[] global = document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "master-styles")!
            .Elements(OdfNamespaces.Draw + "layer-set").Select(element => element.ToString()).ToArray();
        string[] artwork = page.Master!.Elements().Select(element => element.ToString()).ToArray();
        string shapeXml = page.Element.ToString();
        foreach (var layer in document.Layers) page.MasterLayers.Add(layer.Name, layer.Display).IsProtected = layer.IsProtected;
        page.MasterLayers.Find("V--")!.Display = OdgLayerDisplay.Printer;
        foreach (var read in RoundTrips(document)) {
            Assert.True(XNode.DeepEquals(settings, read.GetXml("settings.xml").Root!.Element(OdfNamespaces.Office + "settings")));
            Assert.Equal(global, read.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "master-styles")!
                .Elements(OdfNamespaces.Draw + "layer-set").Select(element => element.ToString()).ToArray());
            Assert.Equal(artwork, read.Pages[0].Master!.Elements().Where(element => element.Name != OdfNamespaces.Draw + "layer-set")
                .Select(element => element.ToString()).ToArray());
            Assert.Equal(shapeXml, read.Pages[0].Element.ToString());
            Assert.Equal(OdgLayerDisplay.Screen, read.Layers.Find("V--")!.Display);
            Assert.Equal(OdgLayerDisplay.Printer, read.Pages[0].MasterLayers.Find("V--")!.Display);
            var screen = read.Pages[0].ToDrawing().Value.Elements.OfType<OfficeDrawingRichText>().Select(t => t.PlainText).ToArray();
            Assert.DoesNotContain(screen, text => text.Contains("V--"));
            Assert.Contains(screen, text => text.Contains("V-L"));
            var print = read.Pages[0].ToDrawing(forPrint: true).Value.Elements.OfType<OfficeDrawingRichText>().Select(t => t.PlainText).ToArray();
            Assert.Contains(print, text => text.Contains("V--"));
            Assert.DoesNotContain(print, text => text.Contains("V-L"));
        }
    }

    private static void AssertLabels(OdgPage page, string? screen, string? print) {
        foreach (var intent in new[] { (ForPrint: false, Label: screen), (ForPrint: true, Label: print) }) {
            var text = page.ToDrawing(forPrint: intent.ForPrint).Value.Elements.OfType<OfficeDrawingRichText>().ToArray();
            Assert.Equal(intent.Label == null ? 0 : 1, text.Length);
            if (intent.Label != null) Assert.Contains(">" + intent.Label + "<", OfficeDrawingSvgExporter.ToSvg(page.ToDrawing(forPrint: intent.ForPrint).Value));
        }
    }
    private static string[] Parts(OdgDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString() };
    private static OdgDocument[] RoundTrips(OdgDocument document) {
        byte[] package = document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource });
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { document, OdgDocument.Load(new MemoryStream(package)), OdgDocument.LoadFlatXml(flat) };
    }
}
