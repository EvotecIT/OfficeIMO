using System.IO;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentFlatXmlSettingsTests {
    [Fact]
    public void MissingFlatSettingsRemainAbsentInsteadOfBecomingAnInvalidEmptyContainer() {
        OdfDocument[] documents = { OdtDocument.Create(), OdsDocument.Create(), OdpPresentation.Create(), OdgDocument.Create() };
        foreach (OdfDocument document in documents) {
            using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
            var imported = OdfDocument.LoadFlatXml(flat);
            Assert.Null(imported.GetXml("settings.xml").Root!.Element(OdfNamespaces.Office + "settings"));
            var package = OdfDocument.Load(new MemoryStream(imported.ToBytes()));
            Assert.Null(package.GetXml("settings.xml").Root!.Element(OdfNamespaces.Office + "settings"));
        }
    }

    [Fact]
    public void NativeFlatConfigurationRemainsPreserved() {
        var document = OdgDocument.Create();
        XNamespace config = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
        var settings = new XElement(OdfNamespaces.Office + "settings", new XElement(config + "config-item-set",
            new XAttribute(config + "name", "Views"), new XElement(config + "config-item", new XAttribute(config + "name", "GridVisible"), new XAttribute(config + "type", "boolean"), "true")));
        document.GetXml("settings.xml").Root!.Add(settings); document.MarkPartDirty("settings.xml");
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        var imported = OdfDocument.LoadFlatXml(flat);
        Assert.True(XNode.DeepEquals(settings, imported.GetXml("settings.xml").Root!.Element(OdfNamespaces.Office + "settings")));
    }
}
