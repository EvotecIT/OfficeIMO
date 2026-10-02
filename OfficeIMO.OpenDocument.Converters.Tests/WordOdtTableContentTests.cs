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
