using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentRepeatedTableRowTests {
    [Fact]
    public void TextAndPresentationTablesExposeRepeatedRowsLogically() {
        using OdtDocument text = OdtDocument.Create();
        OdtTable textTable = text.AddTable(1, 1, "RepeatedText");
        textTable.Cell(0, 0).Text = "Text row";
        textTable.Element.Elements(OdfNamespaces.Table + "table-row").Single()
            .SetAttributeValue(OdfNamespaces.Table + "number-rows-repeated", 3);

        using OdpPresentation presentation = OdpPresentation.Create();
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

        using OdtDocument reopenedText = OdtDocument.Open(new MemoryStream(text.ToBytes()));
        using OdpPresentation reopenedPresentation = OdpPresentation.Open(new MemoryStream(presentation.ToBytes()));
        Assert.Equal(new[] { "Text row", "Changed text row", "Text row" },
            reopenedText.Tables.Single().Rows.Select(row => row.Cells[0].Text));
        Assert.Equal(new[] { "Slide row", "Changed slide row", "Slide row" }, reopenedPresentation.Slides.Single()
            .Shapes.OfType<OdpTable>().Single().Rows.Select(row => row.Cells[0].Text));
    }
}
