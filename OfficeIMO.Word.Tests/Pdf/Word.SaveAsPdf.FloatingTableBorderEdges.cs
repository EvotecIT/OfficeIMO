using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    // A three-row floating table; positions measured from Word's own PDF export.
    // The first row's outer left border places the grid, and text clears the widest
    // outer border on each side across all rows.
    [Theory]
    [InlineData("grid6", "left", 190.61, 78.14)]
    [InlineData("grid6", "right", 72.02, 436.75)]
    [InlineData("style", "left", 191.33, 78.62)]
    [InlineData("row2-thick", "left", 191.57, 77.66)]
    [InlineData("row2-thick", "right", 72.02, 435.19)]
    [InlineData("row1-thick", "left", 193.61, 79.70)]
    [InlineData("row3-cells-only", "left", 190.61, 77.42)]
    public void SaveAsPdf_FloatingTableEdgesCountEveryOuterBorder(string borders, string alignment, double expectedText, double expectedCell) {
        using WordDocument document = CreateFloatingEdgeDocument();
        string tableBorders = borders switch {
            "grid6" => BorderXml("tblBorders", 12, "top", "left", "bottom", "right", "insideH", "insideV"),
            "row2-thick" or "row1-thick" => BorderXml("tblBorders", 4, "top", "left", "bottom", "right"),
            _ => string.Empty
        };
        string RowBorders(int row) => (borders, row) switch {
            ("row2-thick", 1) or ("row1-thick", 0) => BorderXml("tcBorders", 36, "left", "right"),
            ("row3-cells-only", 2) => BorderXml("tcBorders", 24, "left", "right"),
            _ => string.Empty
        };
        string style = borders == "style" ? "<w:tblStyle w:val=\"FloatEdgeGrid\"/>" : string.Empty;
        if (borders == "style") {
            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(new Style(
                "<w:style xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:type=\"table\" w:styleId=\"FloatEdgeGrid\">" +
                "<w:name w:val=\"Float Edge Grid\"/><w:tblPr>" +
                BorderXml("tblBorders", 18, "top", "left", "bottom", "right", "insideH", "insideV") + "</w:tblPr></w:style>"));
        }
        string position = alignment == "right" ? "w:tblpXSpec=\"right\"" : "w:tblpXSpec=\"left\"";
        string rows = string.Concat(new[] { "ROWONE", "ROWTWO", "ROWTHREE" }.Select((text, row) =>
            $"<w:tr><w:tc><w:tcPr><w:tcW w:w=\"2160\" w:type=\"dxa\"/>{RowBorders(row)}</w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc></w:tr>"));
        var table = new Table("<w:tbl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:tblPr>" + style +
            $"<w:tblpPr w:leftFromText=\"180\" w:rightFromText=\"180\" w:vertAnchor=\"text\" w:horzAnchor=\"margin\" {position} w:tblpY=\"1\"/>" +
            "<w:tblW w:w=\"2160\" w:type=\"dxa\"/>" + tableBorders + "</w:tblPr><w:tblGrid><w:gridCol w:w=\"2160\"/></w:tblGrid>" + rows + "</w:tbl>");
        Body body = document._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
        body.InsertBefore(table, body.Elements<SectionProperties>().FirstOrDefault());
        document.AddParagraph(FloatingEdgeWords(60));

        using PdfPigDocument pdf = RenderFloatingEdgePdf(document);
        Assert.Equal(expectedText, FindFloatingEdgeWord(pdf, "w01").BoundingBox.Left, 0.15);
        Assert.Equal(expectedCell, FindFloatingEdgeWord(pdf, "ROWONE").BoundingBox.Left, 0.15);
    }

    [Theory]
    [InlineData("left", 189.74, 77.42)]
    [InlineData("right", 72.02, 436.75)]
    public void SaveAsPdf_InteriorCellBorderIsNotAFloatingTableOutsideEdge(string alignment, double expectedText, double expectedCell) {
        // Word treats a right border on the first of two cells as an interior border: the table
        // keeps the borderless outside edge.
        using WordDocument document = CreateFloatingEdgeDocument();
        string position = alignment == "right" ? "w:tblpXSpec=\"right\"" : "w:tblpXSpec=\"left\"";
        var table = new Table("<w:tbl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:tblPr>" +
            $"<w:tblpPr w:leftFromText=\"180\" w:rightFromText=\"180\" w:vertAnchor=\"text\" w:horzAnchor=\"margin\" {position} w:tblpY=\"1\"/>" +
            "<w:tblW w:w=\"2160\" w:type=\"dxa\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"1080\"/><w:gridCol w:w=\"1080\"/></w:tblGrid><w:tr>" +
            "<w:tc><w:tcPr><w:tcW w:w=\"1080\" w:type=\"dxa\"/>" + BorderXml("tcBorders", 48, "right") + "</w:tcPr><w:p><w:r><w:t>CELLA</w:t></w:r></w:p></w:tc>" +
            "<w:tc><w:tcPr><w:tcW w:w=\"1080\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>CELLB</w:t></w:r></w:p></w:tc></w:tr></w:tbl>");
        Body body = document._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
        body.InsertBefore(table, body.Elements<SectionProperties>().FirstOrDefault());
        document.AddParagraph(FloatingEdgeWords(60));

        using PdfPigDocument pdf = RenderFloatingEdgePdf(document);
        Assert.Equal(expectedText, FindFloatingEdgeWord(pdf, "w01").BoundingBox.Left, 0.15);
        Assert.Equal(expectedCell, FindFloatingEdgeWord(pdf, "CELLA").BoundingBox.Left, 0.15);
    }

    private static string BorderXml(string element, int size, params string[] sides) =>
        $"<w:{element}>" + string.Concat(sides.Select(side => $"<w:{side} w:val=\"single\" w:sz=\"{size}\" w:space=\"0\" w:color=\"auto\"/>")) + $"</w:{element}>";
}
