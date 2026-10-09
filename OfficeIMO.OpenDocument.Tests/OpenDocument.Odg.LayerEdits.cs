using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgLayerEditTests {
    private static readonly XNamespace Config = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InvalidPrintMaskRejectsVisibilityEditWithoutChangingAnyPart(bool oversized) {
        var document = Fixture();
        Items(document, "PrintableLayers").First().Value = oversized ? new string('A', 131073) : "not-valid-base64";
        string[] before = Parts(document);
        Assert.Throws<InvalidDataException>(() => document.Layers.Find("V--")!.Display = OdgLayerDisplay.Printer);
        Assert.Equal(before, Parts(document));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InvalidPrintMaskRejectsLayerAdditionWithoutLeavingLayerOrContainer(bool existingSet) {
        var document = Fixture();
        Items(document, "PrintableLayers").First().Value = "not-valid-base64";
        if (!existingSet) document.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "layer-set").Remove();
        string[] before = Parts(document);
        Assert.Throws<InvalidDataException>(() => document.Layers.Add("Annotation", OdgLayerDisplay.Printer));
        Assert.Equal(before, Parts(document));
    }

    [Fact]
    public void SynchronizesBothVisibilityMasksInEverySavedViewWithoutChangingOtherLayerBits() {
        var document = Fixture();
        XElement view = Items(document, "VisibleLayers").First().Parent!;
        view.AddAfterSelf(new XElement(view));
        foreach (XElement item in Items(document, "VisibleLayers")) item.Value = Convert.ToBase64String(new byte[] { 0x55 });
        foreach (XElement item in Items(document, "PrintableLayers")) item.Value = Convert.ToBase64String(new byte[] { 0xAA });
        XElement layer = document.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "layer")
            .Single(element => (string?)element.Attribute(OdfNamespaces.Draw + "name") == "V--");
        int index = layer.ElementsBeforeSelf(OdfNamespaces.Draw + "layer").Count();
        document.Layers.Find("V--")!.Display = OdgLayerDisplay.Printer;
        foreach (var read in RoundTrips(document)) {
            Assert.Equal(OdgLayerDisplay.Printer, read.Layers.Find("V--")!.Display);
            Assert.All(Items(read, "VisibleLayers"), item => Assert.Equal(new byte[] { (byte)(0x55 & ~(1 << index)) }, Convert.FromBase64String(item.Value)));
            Assert.All(Items(read, "PrintableLayers"), item => Assert.Equal(new byte[] { (byte)(0xAA | (1 << index)) }, Convert.FromBase64String(item.Value)));
            foreach (XElement element in read.GetXml("styles.xml").Descendants(OdfNamespaces.Draw + "layer")) element.Attribute(OdfNamespaces.Draw + "display")?.Remove();
            Assert.Equal(OdgLayerDisplay.Printer, read.Layers.Find("V--")!.Display);
        }
    }

    private static OdgDocument Fixture() => OdgDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-layers.odg"));
    private static XElement[] Items(OdgDocument document, string name) => document.GetXml("settings.xml").Descendants(Config + "config-item")
        .Where(item => (string?)item.Attribute(Config + "name") == name).ToArray();
    private static string[] Parts(OdgDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString(), document.GetXml("settings.xml").ToString() };
    private static OdgDocument[] RoundTrips(OdgDocument document) {
        using var stream = new MemoryStream(); document.SaveFlatXml(stream); stream.Position = 0;
        return new[] { document, OdgDocument.Load(new MemoryStream(document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource }))), OdgDocument.LoadFlatXml(stream) };
    }
}
