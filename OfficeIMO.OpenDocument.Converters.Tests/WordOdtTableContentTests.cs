using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class WordOdtTableContentTests {
    [Fact]
    public void ReverseConversionReportsNestedTableLossAndStrictPolicyRejectsIt() {
        using WordDocument source = WordDocument.Create();
        WordTable outer = source.AddTable(1, 1);
        outer.Rows[0].Cells[0].AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "Nested Word text";
        OdfConversionResult<OdtDocument> converted = source.ToOpenDocumentResult();
        Assert.Contains(converted.Report.Mappings, mapping => mapping.Feature == "nested-tables" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void CellListsAndNestedTablesRetainOrderTextAndNumberingUnderStrictConversion() {
        OdtDocument source = OdtDocument.Create();
        OdtTable outer = source.AddTable(1, 1, "Outer");
        XElement cell = outer.Cell(0, 0).Element;
        cell.RemoveNodes();
        cell.Add(new XElement(OdfNamespaces.Text + "p", "Before"),
            new XElement(OdfNamespaces.Text + "list",
                new XElement(OdfNamespaces.Text + "list-item", new XElement(OdfNamespaces.Text + "p", "List item"))));
        OdtTable nested = source.AddTable(1, 1, "Nested");
        nested.Cell(0, 0).Text = "Nested text";
        nested.Element.Remove();
        cell.Add(nested.Element, new XElement(OdfNamespaces.Text + "p", "After"));
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<WordDocument> converted = source.ToWordDocumentResult(new WordOpenDocumentConversionOptions {
            LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss
        });
        using WordDocument result = converted.Value;
        Table output = result.OpenXmlDocument!.MainDocumentPart!.Document!.Body!.Elements<Table>().Single();
        TableCell resultCell = output.Descendants<TableCell>().First();
        Assert.Equal(new[] { "Before", "List item", "Nested text", "After" },
            resultCell.Descendants<Paragraph>().Select(paragraph => paragraph.InnerText).Where(text => text.Length > 0));
        Assert.NotNull(resultCell.Elements<Paragraph>().Single(paragraph => paragraph.InnerText == "List item")
            .ParagraphProperties?.NumberingProperties);
        Assert.Single(resultCell.Elements<Table>());
        Assert.False(converted.Report.HasLoss);
        Assert.Empty(result.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ListItemsHeadersAndAdjacentListBoundariesSurvive(bool inCell) {
        OdtDocument source = OdtDocument.Create();
        OdtList first = source.AddList(true);
        first.AddItem("Item").AddParagraph("Continuation");
        first.AddItem("Next");
        OdtList second = source.AddList(true);
        second.AddItem("Restart");
        XElement[] lists = source.Package.GetXml("content.xml").Descendants(OdfNamespaces.Text + "list").ToArray();
        lists[0].AddFirst(new XElement(OdfNamespaces.Text + "list-header", new XElement(OdfNamespaces.Text + "p", "Header")));
        if (inCell) {
            XElement cell = source.AddTable(1, 1).Cell(0, 0).Element;
            cell.RemoveNodes();
            foreach (XElement list in lists) { list.Remove(); cell.Add(list); }
        }
        source.Package.MarkXmlDirty("content.xml");
        using WordDocument result = source.ToWordDocumentResult(new WordOpenDocumentConversionOptions {
            LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }).Value;
        Paragraph[] paragraphs = result.OpenXmlDocument.MainDocumentPart!.Document!.Body!.Descendants<Paragraph>()
            .Where(paragraph => paragraph.InnerText.Length > 0).ToArray();
        Assert.Equal(new[] { "Header", "Item", "Continuation", "Next", "Restart" }, paragraphs.Select(p => p.InnerText));
        Assert.Null(paragraphs[0].ParagraphProperties?.NumberingProperties);
        Assert.Null(paragraphs[2].ParagraphProperties?.NumberingProperties);
        string[] numberIds = paragraphs.Where(p => p.ParagraphProperties?.NumberingProperties != null)
            .Select(p => p.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value.ToString()).ToArray();
        Assert.Equal(3, numberIds.Length);
        Assert.Equal(numberIds[0], numberIds[1]);
        Assert.NotEqual(numberIds[1], numberIds[2]);
        Assert.Empty(result.ValidateDocument());
    }

    [Fact]
    public void NestedListDoesNotMoveContinuationOrFollowingParentItem() {
        OdtDocument source = OdtDocument.Create();
        OdtList list = source.AddList(true);
        list.AddItem("Parent").AddParagraph("Continuation");
        list.AddItem("Next parent");
        XElement root = source.Package.GetXml("content.xml").Descendants(OdfNamespaces.Text + "list").Single();
        XElement first = root.Elements(OdfNamespaces.Text + "list-item").First();
        first.Elements(OdfNamespaces.Text + "p").First().AddAfterSelf(new XElement(OdfNamespaces.Text + "list",
            new XAttribute(OdfNamespaces.Text + "style-name", (string)root.Attribute(OdfNamespaces.Text + "style-name")!),
            new XElement(OdfNamespaces.Text + "list-item", new XElement(OdfNamespaces.Text + "p", "Nested"))));
        source.Package.MarkXmlDirty("content.xml");
        using WordDocument result = source.ToWordDocumentResult(new WordOpenDocumentConversionOptions {
            LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }).Value;
        Paragraph[] paragraphs = result.OpenXmlDocument.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ToArray();
        Assert.Equal(new[] { "Parent", "Nested", "Continuation", "Next parent" }, paragraphs.Select(p => p.InnerText));
        Assert.Equal(1, paragraphs[1].ParagraphProperties!.NumberingProperties!.NumberingLevelReference!.Val!.Value);
        Assert.Null(paragraphs[2].ParagraphProperties?.NumberingProperties);
        Assert.Equal(paragraphs[0].ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value,
            paragraphs[3].ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value);
    }

    [Fact]
    public void ExplicitListContinuationAndItemRestartsRemainReportedLoss() {
        OdtDocument source = OdtDocument.Create();
        source.AddList(true).AddItem("Item");
        XElement list = source.Package.GetXml("content.xml").Descendants(OdfNamespaces.Text + "list").Single();
        list.SetAttributeValue(OdfNamespaces.Text + "continue-numbering", true);
        list.Elements(OdfNamespaces.Text + "list-item").Single().SetAttributeValue(OdfNamespaces.Text + "start-value", 5);
        source.Package.MarkXmlDirty("content.xml");
        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument result = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "list-numbering" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void TableHeadingBeyondWordOutlineLimitReportsApproximation() {
        OdtDocument source = OdtDocument.Create();
        OdtParagraph paragraph = source.AddTable(1, 1).Cell(0, 0).Paragraphs.Single();
        paragraph.Text = "Deep heading";
        paragraph.HeadingLevel = 10;
        OdfConversionResult<WordDocument> conversion = source.ToWordDocumentResult();
        using WordDocument result = conversion.Value;
        Assert.Contains(conversion.Report.Mappings, mapping => mapping.Feature == "heading-levels" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
        Assert.Throws<OdfConversionLossException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AggregateExpansionIsRejectedBeforeLogicalContentInspection(bool includeImages) {
        OdtDocument source = OdtDocument.Create();
        OdtTable table = source.AddTable(1, 1);
        table.Element.Elements(OdfNamespaces.Table + "table-row").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-rows-repeated", 4096);
        table.Cell(0, 0).Element.SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", 256);
        source.Package.MarkXmlDirty("content.xml");
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { IncludeImages = includeImages }));
        Assert.Contains("MaxConvertedTableCells", error.Message);
    }

    [Fact]
    public void ExpansionBudgetIncludesAllTopLevelTables() {
        OdtDocument source = OdtDocument.Create();
        source.AddTable(1, 2, "First");
        source.AddTable(1, 2, "Second");
        Assert.Throws<InvalidDataException>(() => source.ToWordDocumentResult(
            new WordOpenDocumentConversionOptions { MaxConvertedTableCells = 3 }));
    }

    [Fact]
    public void TextAmplificationIsRejectedEvenWhenCellCountFits() {
        OdtDocument source = OdtDocument.Create();
        OdtTable table = source.AddTable(1, 1);
        table.Cell(0, 0).Text = new string('x', 600);
        table.Element.Elements(OdfNamespaces.Table + "table-row").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-rows-repeated", 200);
        table.Cell(0, 0).Element.SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", 256);
        source.Package.MarkXmlDirty("content.xml");
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => source.ToWordDocumentResult());
        Assert.Contains("MaxConvertedTableTextCharacters", error.Message);
    }
}
