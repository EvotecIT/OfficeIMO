using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OdtAlternateHeaderFooterTests {
    [Fact]
    public void FirstAndLeftHeaderFooterVariantsReopenInSchemaOrder() {
        OdtDocument document = OdtDocument.Create();
        document.AddParagraph("Body");
        Assert.Null(document.PageLayout.FirstHeader);
        Assert.Null(document.PageLayout.LeftFooter);

        document.PageLayout.EnsureFirstFooter().AddParagraph("First footer");
        document.PageLayout.EnsureLeftHeader().AddParagraph("Left header");
        document.PageLayout.EnsureFirstHeader().AddParagraph("First header");
        document.PageLayout.EnsureLeftFooter().AddParagraph("Left footer");
        document.PageLayout.Header.AddParagraph("Default header");
        document.PageLayout.Footer.AddParagraph("Default footer");

        byte[] bytes = document.ToBytes();
        OdtDocument reopened = OdtDocument.Load(new MemoryStream(bytes));
        Assert.Equal("Default header", Assert.Single(reopened.PageLayout.Header.Paragraphs).Text);
        Assert.Equal("Left header", Assert.Single(reopened.PageLayout.LeftHeader!.Paragraphs).Text);
        Assert.Equal("First header", Assert.Single(reopened.PageLayout.FirstHeader!.Paragraphs).Text);
        Assert.Equal("Default footer", Assert.Single(reopened.PageLayout.Footer.Paragraphs).Text);
        Assert.Equal("Left footer", Assert.Single(reopened.PageLayout.LeftFooter!.Paragraphs).Text);
        Assert.Equal("First footer", Assert.Single(reopened.PageLayout.FirstFooter!.Paragraphs).Text);

        XDocument styles = reopened.Package.GetXml("styles.xml");
        XElement master = Assert.Single(styles.Descendants(OdfNamespaces.Style + "master-page"));
        Assert.Equal(new[] { "header", "header-left", "header-first", "footer", "footer-left", "footer-first" },
            master.Elements().Select(element => element.Name.LocalName));
    }
}
