using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, 0, 72)]
    [InlineData(false, 1, 72)]
    [InlineData(false, 0, 144)]
    [InlineData(false, 1, 144)]
    [InlineData(true, 0, 72)]
    [InlineData(true, 1, 72)]
    [InlineData(true, 0, 144)]
    [InlineData(true, 1, 144)]
    public void SaveAsPdf_AutomaticCellPreferenceRespectsPictureRowOrder(bool reversed, int column, int pictureWidth) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutofitPreferenceOrderControl(document);
        WordTableCell cell = table.Rows[1].Cells[column];
        cell.WidthType = WordTableWidthUnit.Auto;
        cell.Width = 0;
        WordParagraph paragraph = cell.Paragraphs[0];
        paragraph.Text = "A";
        using var image = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(2, 1));
        paragraph.AddImage(image, "marker.png", pictureWidth * 96D / 72D, 24);
        paragraph.AddText("B");
        ReverseAutofitPreferenceRows(table, reversed);
        string source = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;

        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        }));
        Assert.Equal(1, pdf.NumberOfPages);
        var page = pdf.GetPage(1);
        var a = Assert.Single(page.Letters, letter => letter.Value == "A");
        var b = Assert.Single(page.Letters, letter => letter.Value == "B");
        var picture = Assert.Single(page.GetImages());
        Assert.Equal(pictureWidth, picture.BoundingBox.Width, 3);
        Assert.Equal(18D, picture.BoundingBox.Height, 3);
        if (reversed) Assert.True(Math.Abs(a.StartBaseLine.Y - b.StartBaseLine.Y) > 1D);
        else {
            Assert.Equal(a.StartBaseLine.Y, b.StartBaseLine.Y, 3);
            Assert.True(a.BoundingBox.Right <= picture.BoundingBox.Left + .5D);
            Assert.True(picture.BoundingBox.Right <= b.BoundingBox.Left + .5D);
        }
        Assert.Contains("M0", page.Text, StringComparison.Ordinal);
        Assert.Contains("M1", page.Text, StringComparison.Ordinal);
        Assert.Equal(source, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    [Theory]
    [InlineData(false, WordTableWidthUnit.Auto)]
    [InlineData(true, WordTableWidthUnit.Auto)]
    [InlineData(false, WordTableWidthUnit.Dxa)]
    [InlineData(true, WordTableWidthUnit.Dxa)]
    public void SaveAsPdf_AutomaticCellPreferenceRespectsBreakableTextRowOrder(bool reversed, WordTableWidthUnit widthType) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutofitPreferenceOrderControl(document);
        WordTableCell cell = table.Rows[1].Cells[0];
        cell.WidthType = widthType;
        cell.Width = widthType == WordTableWidthUnit.Auto ? 0 : 1600;
        cell.Paragraphs[0].Text = "Alpha Bravo Charlie Delta";
        ReverseAutofitPreferenceRows(table, reversed);
        string source = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        }));
        Assert.Equal(1, pdf.NumberOfPages);
        const string expected = "AlphaBravoCharlieDelta";
        var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        int start = Array.FindIndex(letters, letter => letter.Value == "A");
        Assert.True(start >= 0);
        var phrase = letters.Skip(start).Take(expected.Length).ToArray();
        Assert.Equal(expected, string.Concat(phrase.Select(letter => letter.Value)));
        int lines = phrase.Select(letter => Math.Round(letter.StartBaseLine.Y, 3)).Distinct().Count();
        if (!reversed && widthType == WordTableWidthUnit.Auto) Assert.Equal(1, lines);
        else Assert.True(lines > 1);
        Assert.Equal(source, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    private static WordTable CreateAutofitPreferenceOrderControl(WordDocument document) {
        WordTable table = CreateAutomaticWidthControl(document, 1600, 1600);
        table.AddRow();
        for (int row = 0; row < 2; row++)
            for (int column = 0; column < 2; column++) {
                WordTableCell cell = table.Rows[row].Cells[column];
                cell.Width = 1600;
                cell.WidthType = WordTableWidthUnit.Dxa;
                cell.MarginLeftWidth = cell.MarginRightWidth = 0;
                cell.MarginTopWidth = cell.MarginBottomWidth = 0;
                WordParagraph paragraph = cell.Paragraphs[0];
                paragraph.Text = (row == 0 ? "M" : "P") + column;
                paragraph.FontFamily = "Arial";
                paragraph.FontSize = 12;
                paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0;
                paragraph.LineSpacingRule = WordLineSpacingRule.Auto;
                paragraph.LineSpacing = 240;
            }
        table.GridColumnWidth = new List<int> { 1600, 1600 };
        return table;
    }

    private static void ReverseAutofitPreferenceRows(WordTable table, bool reverse) {
        if (!reverse) return;
        W.TableRow[] rows = table._table.Elements<W.TableRow>().ToArray();
        rows[1].Remove();
        table._table.InsertBefore(rows[1], rows[0]);
    }
}
