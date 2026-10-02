using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentFlatStyleScopeTests {
    [Fact]
    public void SharedFlatAutomaticStyleIsAvailableToBothBodyAndHeader() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph body = source.AddParagraph("Body");
        body.Bold = true;
        source.PageLayout.Header.AddParagraph("Header").StyleName = body.StyleName;
        XDocument flat = source.ToFlatXml();
        using var stream = new MemoryStream();
        flat.Save(stream);
        stream.Position = 0;
        OdtDocument loaded = OdtDocument.LoadFlatXml(stream);
        Assert.True(loaded.ContentBlocks.Single().Paragraph!.Bold);
        Assert.True(loaded.PageLayout.Header.Paragraphs.Single().Bold);
        Assert.True(loaded.Validate().IsValid);
    }

    [Fact]
    public void ConflictingPartLocalNamesRemainDistinctAfterFlatRoundTrip() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph body = source.AddParagraph("Body");
        body.Bold = true;
        OdtParagraph header = source.PageLayout.Header.AddParagraph("Header");
        header.Italic = true;
        XElement headerStyle = source.Styles.FindInPart(OdfStyleFamily.Paragraph, header.StyleName!, "styles.xml")!.Element;
        headerStyle.SetAttributeValue(OdfNamespaces.Style + "name", body.StyleName);
        header.StyleName = body.StyleName;
        source.Package.MarkXmlDirty("styles.xml");
        XDocument flat = source.ToFlatXml();
        string[] names = flat.Root!.Element(OdfNamespaces.Office + "automatic-styles")!
            .Elements(OdfNamespaces.Style + "style").Select(element => (string)element.Attribute(OdfNamespaces.Style + "name")!).ToArray();
        Assert.Equal(names.Length, names.Distinct().Count());
        using var stream = new MemoryStream();
        flat.Save(stream);
        stream.Position = 0;
        OdtDocument loaded = OdtDocument.LoadFlatXml(stream);
        Assert.True(loaded.ContentBlocks.Single().Paragraph!.Bold);
        Assert.Null(loaded.ContentBlocks.Single().Paragraph!.Italic);
        Assert.True(loaded.PageLayout.Header.Paragraphs.Single().Italic);
        Assert.Null(loaded.PageLayout.Header.Paragraphs.Single().Bold);
        Assert.True(loaded.Validate().IsValid);
    }

    [Fact]
    public void SparseReadIndexesTrackNativeAndMarkedXmlEditsWithoutExpandingRuns() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        sheet.Cell(1000, 1000).SetString("last");
        Assert.Equal("last", sheet.GetValue(1000, 1000).ToString());
        sheet.Cell(500, 500).SetString("middle");
        Assert.Equal("middle", sheet.GetValue(500, 500).ToString());
        Assert.Equal("last", sheet.GetValue(1000, 1000).ToString());
        sheet.Cell(500, 500).Formula = "of:=1+1";
        Assert.Equal("of:=1+1", sheet.GetFormula(500, 500));
        XElement firstRow = sheet.Element.Elements(OdfNamespaces.Table + "table-row").First();
        firstRow.SetAttributeValue(OdfNamespaces.Table + "number-rows-repeated", OdsRepeatModel.Read(firstRow, OdfNamespaces.Table + "number-rows-repeated") + 1);
        source.Package.MarkXmlDirty("content.xml");
        Assert.Equal("of:=1+1", sheet.GetFormula(501, 500));
        sheet.Cell(502, 500).SetString("after shift");
        Assert.Equal("of:=1+1", sheet.GetFormula(501, 500));
        Assert.Equal("after shift", sheet.GetValue(502, 500).ToString());
        Assert.True(sheet.RowRuns.Count < 10);
    }
}
