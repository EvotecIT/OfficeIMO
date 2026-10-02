using System.Linq;
using System.Xml.Linq;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentStyleLookupTests {
    [Fact]
    public void ParagraphsUseFamilyDefaultsAfterExplicitStyleProperties() {
        OdtDocument document = OdtDocument.Create();
        XElement styles = document.Package.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!;
        styles.Add(new XElement(OdfNamespaces.Style + "default-style",
            new XAttribute(OdfNamespaces.Style + "family", "paragraph"),
            new XElement(OdfNamespaces.Style + "text-properties",
                new XAttribute(OdfNamespaces.Fo + "font-weight", "bold"),
                new XAttribute(OdfNamespaces.Fo + "font-size", "20pt"))));
        document.Package.MarkXmlDirty("styles.xml");
        OdtParagraph unstyled = document.AddParagraph("Default");
        OdtParagraph styled = document.AddParagraph("Explicit");
        styled.StyleName = document.Styles.CreateNamed("Italic", OdfStyleFamily.Paragraph).Name;
        styled.Italic = true;
        Assert.True(unstyled.Bold);
        Assert.Equal(OdfLength.Parse("20pt"), unstyled.FontSize);
        Assert.True(styled.Bold);
        Assert.True(styled.Italic);
        styled.Bold = false;
        Assert.False(styled.Bold);
    }

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
