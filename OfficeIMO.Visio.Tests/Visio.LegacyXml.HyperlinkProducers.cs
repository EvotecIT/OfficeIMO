using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioLegacyHyperlinkProducerTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static string Fixture => Path.Combine(AppContext.BaseDirectory, "Fixtures", "LegacyXml", "vdxtosvg-basicshapes.vdx");

    [Fact]
    public void IndependentVisio2002HyperlinkKeepsNamedIdentityAndEverySourceCell() {
        XDocument original = XDocument.Load(Fixture);
        XElement source = original.Descendants(original.Root!.Name.Namespace + "Hyperlink").Single();
        var document = VisioDocument.LoadLegacyXml(Fixture).Value;
        VisioHyperlink link = document.Pages[0].Shapes.Single(s => s.Id == "13").Hyperlinks.Single();
        Assert.Equal("Row_21", link.RowName);
        Assert.Equal("VDXtoSVG Home Page", link.Description);
        Assert.Equal("http://vdxtosvg.sourceforge.net/", link.Address);
        Assert.Equal("\ue000", link.SubAddress);
        Assert.Equal("\ue000", link.ExtraInfo);

        foreach (VisioDocument reopened in Reopenings(document)) {
            XElement row = Export(reopened).Descendants(Legacy + "Hyperlink").Single();
            AssertSourceRow(source, row);
        }
    }

    [Fact]
    public void IndependentHyperlinkCopyCanBeEditedWithoutChangingTheSourceOrNeighboringCells() {
        XDocument original = XDocument.Load(Fixture);
        XElement sourceRow = original.Descendants(original.Root!.Name.Namespace + "Hyperlink").Single();
        var document = VisioDocument.LoadLegacyXml(Fixture).Value;
        VisioPage source = document.Pages[0];
        VisioPage copy = document.DuplicatePage(source, "Hyperlink copy");
        VisioShape copiedShape = copy.Shapes.Single(s => s.Hyperlinks.Count != 0);
        VisioHyperlink edited = copiedShape.Hyperlinks.Single();
        edited.Address = "https://example.org/edited";
        edited.Frame = edited.Frame;
        copiedShape.AddPageHyperlink(source.Name, "Return to source");

        foreach (VisioDocument reopened in Reopenings(document)) {
            XDocument xml = Export(reopened);
            XElement originalRow = xml.Descendants(Legacy + "Page").First().Descendants(Legacy + "Hyperlink").Single();
            AssertSourceRow(sourceRow, originalRow);
            XElement copiedRow = xml.Descendants(Legacy + "Page").Last().Descendants(Legacy + "Hyperlink")
                .Single(row => (string?)row.Attribute("NameU") == "Row_21");
            Assert.Equal("1", (string?)copiedRow.Attribute("ID"));
            Assert.Equal("https://example.org/edited", copiedRow.Element(Legacy + "Address")!.Value);
            Assert.Null(copiedRow.Element(Legacy + "Frame")!.Attribute("F"));
            foreach (string name in new[] { "Description", "SubAddress", "ExtraInfo", "NewWindow", "Default" })
                AssertCell(sourceRow.Element(sourceRow.Name.Namespace + name)!, copiedRow.Element(Legacy + name)!);
            VisioShape reloadedCopy = reopened.Pages.Last().Shapes.Single(s => s.Hyperlinks.Count != 0);
            Assert.Equal(2, reloadedCopy.Hyperlinks.Count);
            Assert.Equal(source.Name, reloadedCopy.Hyperlinks.Single(l => l.Description == "Return to source").SubAddress);
        }
        Assert.Equal("http://vdxtosvg.sourceforge.net/", source.Shapes.Single(s => s.Hyperlinks.Count != 0).Hyperlinks.Single().Address);
        Assert.Null(copiedShape.Hyperlinks.Last().RowName);
    }

    private static void AssertSourceRow(XElement source, XElement row) {
        foreach (XAttribute attribute in source.Attributes().Where(a => !a.IsNamespaceDeclaration))
            Assert.Equal(attribute.Value, (string?)row.Attribute(attribute.Name));
        foreach (XElement cell in source.Elements()) AssertCell(cell, row.Element(Legacy + cell.Name.LocalName)!);
    }

    private static void AssertCell(XElement source, XElement cell) {
        Assert.Equal(source.Value, cell.Value);
        Assert.Equal(source.Attributes().Where(a => !a.IsNamespaceDeclaration).Select(a => a.Name + "=" + a.Value).OrderBy(a => a),
            cell.Attributes().Where(a => !a.IsNamespaceDeclaration).Select(a => a.Name + "=" + a.Value).OrderBy(a => a));
    }

    private static IEnumerable<VisioDocument> Reopenings(VisioDocument document) {
        yield return document;
        VisioDocument package = VisioDocument.Load(new MemoryStream(document.ToBytes()));
        yield return package;
        yield return VisioDocument.LoadLegacyXml(new MemoryStream(package.ToLegacyXmlResult().Value)).Value;
    }

    private static XDocument Export(VisioDocument document) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value));
}
