using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentRepeatedTableRowTests {
    [Theory]
    [InlineData("text")]
    [InlineData("bold")]
    [InlineData("span")]
    [InlineData("nested-link")]
    [InlineData("field")]
    public void EditingAParagraphOrItsChildrenChangesOnlyTheLogicalTextCell(string edit) {
        OdtDocument document = OdtDocument.Create();
        OdtTable table = document.AddTable(1, 1, "Repeated");
        OdtParagraph paragraph = table.Cell(0, 0).Paragraphs.Single();
        paragraph.Text = "original";
        paragraph.AddSpan("span").AddHyperlink("link", "https://example.com");
        paragraph.AddField(OdtFieldKind.Date, "date");
        table.Element.Descendants(OdfNamespaces.Table + "table-cell").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", 3);
        table.Element.Elements(OdfNamespaces.Table + "table-row").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-rows-repeated", 3);
        document.Package.MarkXmlDirty("content.xml");
        string before = table.Element.ToString();
        OdtParagraph middle = table.Cell(1, 1).Paragraphs.Single();
        _ = middle.InlineNodes;
        _ = middle.Spans;
        Assert.Equal(before, table.Element.ToString());
        if (edit == "text") middle.Text = "changed";
        else if (edit == "bold") middle.Bold = true;
        else if (edit == "span") middle.Spans[0].Text = "changed";
        else if (edit == "nested-link") middle.InlineNodes.Single(node => node.Span != null).Children.Single(node => node.Hyperlink != null).Hyperlink!.Href = "https://changed.example";
        else middle.Fields.Single().DisplayText = "changed";
        OdtDocument reopened = OdtDocument.Load(new MemoryStream(document.ToBytes()));
        OdtTable result = reopened.Tables.Single();
        for (int row = 0; row < 3; row++) for (int column = 0; column < 3; column++) {
            OdtParagraph actual = result.Cell(row, column).Paragraphs.Single();
            bool changed = row == 1 && column == 1;
            if (edit == "bold") Assert.Equal(changed ? true : (bool?)null, actual.Bold);
            else if (edit == "nested-link") Assert.Equal(changed ? "https://changed.example" : "https://example.com", actual.Hyperlinks.Single().Href);
            else Assert.Equal(changed, actual.Text.Contains("changed"));
        }
    }

    [Fact]
    public void ReadingPresentationParagraphsStaysSparseAndRunEditsAreLogical() {
        OdpPresentation document = OdpPresentation.Create();
        OdpTable table = document.AddSlide("Repeated").AddTable(OdfRect.FromCentimeters(1, 1, 8, 4), 1, 1);
        table.Cell(0, 0).Paragraphs.Single().AddRun("original");
        table.Element.Descendants(OdfNamespaces.Table + "table-cell").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", 3);
        document.Package.MarkXmlDirty("content.xml");
        string before = table.Element.ToString();
        OdpParagraph middle = table.Cell(0, 1).Paragraphs.Single();
        _ = middle.InlineNodes;
        Assert.Equal(before, table.Element.ToString());
        middle.Runs.Single().Text = "changed";
        OdpTable reopened = OdpPresentation.Load(new MemoryStream(document.ToBytes())).Slides.Single().Shapes.OfType<OdpTable>().Single();
        Assert.Equal(new[] { "original", "changed", "original" }, reopened.Rows[0].Cells.Select(cell => cell.Text));
    }

    [Fact]
    public void TextAndPresentationTablesRejectExcessiveLogicalRepeats() {
        OdtDocument text = OdtDocument.Create();
        OdtTable textTable = text.AddTable(1, 1, "BoundedText");
        textTable.Element.Elements(OdfNamespaces.Table + "table-row").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-rows-repeated", 1000001);

        OdpPresentation presentation = OdpPresentation.Create();
        OdpTable presentationTable = presentation.AddSlide("Bounded").AddTable(
            OdfRect.FromCentimeters(1, 1, 8, 4), 1, 1, "BoundedPresentation");
        presentationTable.Element.Descendants(OdfNamespaces.Table + "table-row").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-rows-repeated", 1000001);

        Assert.Throws<InvalidDataException>(() => _ = textTable.Rows);
        Assert.Throws<InvalidDataException>(() => _ = presentationTable.Rows);
    }

    [Fact]
    public void TextAndPresentationTablesExposeRepeatedRowsLogically() {
        OdtDocument text = OdtDocument.Create();
        OdtTable textTable = text.AddTable(1, 1, "RepeatedText");
        textTable.Cell(0, 0).Text = "Text row";
        textTable.Element.Elements(OdfNamespaces.Table + "table-row").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-rows-repeated", 3);

        OdpPresentation presentation = OdpPresentation.Create();
        OdpTable presentationTable = presentation.AddSlide("Repeated").AddTable(
            OdfRect.FromCentimeters(1, 1, 8, 4), 1, 1, "RepeatedPresentation");
        presentationTable.Cell(0, 0).Text = "Slide row";
        presentationTable.Element.Descendants(OdfNamespaces.Table + "table-row").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-rows-repeated", 3);

        Assert.Equal(new[] { "Text row", "Text row", "Text row" },
            textTable.Rows.Select(row => row.Cells[0].Text));
        Assert.Equal(new[] { "Slide row", "Slide row", "Slide row" },
            presentationTable.Rows.Select(row => row.Cells[0].Text));

        textTable.Cell(1, 0).Text = "Changed text row";
        presentationTable.Cell(1, 0).Text = "Changed slide row";

        Assert.Equal(new[] { "Text row", "Changed text row", "Text row" },
            textTable.Rows.Select(row => row.Cells[0].Text));
        Assert.Equal(new[] { "Slide row", "Changed slide row", "Slide row" },
            presentationTable.Rows.Select(row => row.Cells[0].Text));

        OdtDocument reopenedText = OdtDocument.Load(new MemoryStream(text.ToBytes()));
        OdpPresentation reopenedPresentation = OdpPresentation.Load(new MemoryStream(presentation.ToBytes()));
        Assert.Equal(new[] { "Text row", "Changed text row", "Text row" },
            reopenedText.Tables.Single().Rows.Select(row => row.Cells[0].Text));
        Assert.Equal(new[] { "Slide row", "Changed slide row", "Slide row" }, reopenedPresentation.Slides.Single()
            .Shapes.OfType<OdpTable>().Single().Rows.Select(row => row.Cells[0].Text));
    }

    [Fact]
    public void CachedTextAndPresentationCellCollectionsTrackRepeatedCellSplits() {
        OdtDocument text = OdtDocument.Create();
        OdtTable textTable = text.AddTable(1, 1, "RepeatedTextCells");
        textTable.Cell(0, 0).Text = "Original";
        textTable.Element.Descendants(OdfNamespaces.Table + "table-cell").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", 3);
        IReadOnlyList<OdtTableCell> textCells = textTable.Rows[0].Cells;

        OdpPresentation presentation = OdpPresentation.Create();
        OdpTable presentationTable = presentation.AddSlide("RepeatedCells").AddTable(
            OdfRect.FromCentimeters(1, 1, 8, 4), 1, 1, "RepeatedPresentationCells");
        presentationTable.Cell(0, 0).Text = "Original";
        presentationTable.Element.Descendants(OdfNamespaces.Table + "table-cell").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-columns-repeated", 3);
        IReadOnlyList<OdpTableCell> presentationCells = presentationTable.Rows[0].Cells;

        textCells[1].Text = "Middle";
        textCells[2].Text = "Last";
        presentationCells[1].Text = "Middle";
        presentationCells[2].Text = "Last";

        Assert.Equal(new[] { "Original", "Middle", "Last" }, textTable.Rows[0].Cells.Select(cell => cell.Text));
        Assert.Equal(new[] { "Original", "Middle", "Last" }, presentationTable.Rows[0].Cells.Select(cell => cell.Text));
    }
}
