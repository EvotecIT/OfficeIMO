using System.Linq;
using System.Xml.Linq;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentStyleLookupTests {
    [Fact]
    public void IndexedLookupKeepsPartPrecedenceAndTracksMarkedXmlChanges() {
        OdpPresentation document = OdpPresentation.Create();
        OdfStyle named = document.Styles.CreateNamed("Shared", OdfStyleFamily.Paragraph);
        Assert.Same(named.Element, document.Styles.Find(OdfStyleFamily.Paragraph, "Shared")!.Element);

        XElement automaticStyles = document.Package.GetXml("content.xml").Root!
            .Element(OdfNamespaces.Office + "automatic-styles")!;
        var automatic = new XElement(OdfNamespaces.Style + "style",
            new XAttribute(OdfNamespaces.Style + "name", "Shared"),
            new XAttribute(OdfNamespaces.Style + "family", "paragraph"));
        automaticStyles.Add(automatic);
        document.Package.MarkXmlDirty("content.xml");

        Assert.Same(automatic, document.Styles.Find(OdfStyleFamily.Paragraph, "Shared")!.Element);
        Assert.Same(automatic, document.Styles.FindInPart(OdfStyleFamily.Paragraph, "Shared", "content.xml")!.Element);
        Assert.Same(named.Element, document.Styles.FindInPart(OdfStyleFamily.Paragraph, "Shared", "styles.xml")!.Element);

        automatic.Remove();
        document.Package.MarkXmlDirty("content.xml");
        Assert.Same(named.Element, document.Styles.Find(OdfStyleFamily.Paragraph, "Shared")!.Element);
        Assert.Null(document.Styles.Find(OdfStyleFamily.Paragraph, "Absent"));
    }
}
