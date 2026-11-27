using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    private const string TableExtensionWordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static OpenXmlUnknownElement AddTableExtension(OpenXmlElement parent, string innerXml) {
        var extension = new OpenXmlUnknownElement("x", "tableExtension", "urn:officeimo:test");
        extension.InnerXml = innerXml;
        parent.AppendChild(extension);
        return extension;
    }

    [Fact]
    public void TableAndCellSettingsInsideExtensionMarkupAreNotTheTables() {
        using var document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.AddTableCellProperties();
        string w = $"xmlns:w=\"{TableExtensionWordNamespace}\"";
        OpenXmlUnknownElement cellExtension = AddTableExtension(cell._tableCellProperties!,
            $"<w:tcBorders {w}><w:left w:val=\"double\" w:sz=\"24\" w:color=\"FF0000\"/></w:tcBorders>");
        OpenXmlUnknownElement tableExtension = AddTableExtension(table._tableProperties!,
            $"<w:tblBorders {w}><w:top w:val=\"double\" w:sz=\"24\"/></w:tblBorders>" +
            $"<w:tblCellMar {w}><w:top w:w=\"500\" w:type=\"dxa\"/></w:tblCellMar>" +
            $"<w:shd {w} w:val=\"pct10\" w:fill=\"00FF00\"/>");
        string cellExtensionMarkup = cellExtension.OuterXml;
        string tableExtensionMarkup = tableExtension.OuterXml;

        // Settings that only exist inside the extensions read as unset.
        Assert.Null(cell.Borders.LeftStyle);
        Assert.Null(cell.Borders.LeftColor);
        WordTableStyleDetails details = table.StyleDetails!;
        Assert.Null(details.GetBorderProperties(WordTableBorderSide.Top).Style);
        Assert.Null(details.MarginDefaultTop);
        Assert.Null(details.ShadingPattern);

        // Setting them writes the table's and cell's own elements and leaves the extensions alone.
        cell.Borders.LeftStyle = WordBorderStyle.Single;
        details.SetCustomBorders(topStyle: WordBorderStyle.Dotted);
        details.MarginDefaultTop = OfficeIMO.OfficeLength.Points(5);
        details.ShadingFillColor = OfficeIMO.OfficeColor.Parse("#0000FF");
        Assert.Equal(BorderValues.Single, cell._tableCellProperties!.TableCellBorders!.LeftBorder!.Val!.Value);
        Assert.Equal(BorderValues.Dotted, table._tableProperties!.TableBorders!.TopBorder!.Val!.Value);
        Assert.Equal("100", table._tableProperties.TableCellMarginDefault!.TopMargin!.Width!.Value);
        Assert.Equal("0000FF", table._tableProperties.GetFirstChild<Shading>()!.Fill!.Value);
        Assert.Equal(cellExtensionMarkup, cellExtension.OuterXml);
        Assert.Equal(tableExtensionMarkup, tableExtension.OuterXml);

        // Applying the table's borders copies only its own top border to the cell.
        details.ApplyBordersToAllCells();
        Assert.Equal(WordBorderStyle.Dotted, cell.Borders.TopStyle);
        Assert.Null(cell._tableCellProperties.TableCellBorders.BottomBorder);
    }

    [Fact]
    public void RejectedCellBorderEditsLeaveTheCellAsItWas() {
        using var document = WordDocument.Create();
        WordTableCell cell = document.AddTable(1, 1).Rows[0].Cells[0];

        Assert.ThrowsAny<ArgumentException>(() => cell.Borders.LeftStyle = (WordBorderStyle)(-1));
        Assert.ThrowsAny<ArgumentException>(() => cell.Borders.LeftSize = OfficeIMO.OfficeLength.Points(-1));
        Assert.Null(cell._tableCellProperties?.TableCellBorders);

        // Both a w:left and a w:start can be set; w:left comes first, as Word writes it.
        cell.Borders.StartStyle = WordBorderStyle.Single;
        cell.Borders.LeftStyle = WordBorderStyle.Dotted;
        Assert.Equal(new[] { "left", "start" }, cell._tableCellProperties!.TableCellBorders!.ChildElements.Select(edge => edge.LocalName));
    }
}
