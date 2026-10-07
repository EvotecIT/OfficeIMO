using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordTableLayoutMode.Fixed, WordTableWidthUnit.Dxa)]
    [InlineData(WordTableLayoutMode.Fixed, WordTableWidthUnit.Auto)]
    [InlineData(WordTableLayoutMode.Fixed, WordTableWidthUnit.Pct)]
    [InlineData(WordTableLayoutMode.AutoFit, WordTableWidthUnit.Dxa)]
    public void SaveAsPdf_NoWrapRetainsWrappingForFixedLayoutOrAbsoluteCellWidth(
        WordTableLayoutMode layout, WordTableWidthUnit widthType) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        table.ConditionalFormattingFirstRow = false;
        table.Width = 3000;
        table.WidthType = WordTableWidthUnit.Dxa;
        table.LayoutMode = layout;
        table.GridColumnWidth = new List<int> { 3000 };
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.Width = widthType == WordTableWidthUnit.Pct ? 5000 : 3000;
        cell.WidthType = widthType;
        cell.WrapText = false;
        cell.MarginLeftWidth = 108;
        cell.MarginRightWidth = 108;
        cell.MarginTopWidth = 0;
        cell.MarginBottomWidth = 0;
        WordParagraph paragraph = cell.Paragraphs[0];
        const string text = "Start Alpha Beta Gamma Delta Epsilon Zeta Eta Theta";
        paragraph.Text = text;
        paragraph.FontFamily = "Arial";
        paragraph.FontSize = 12;
        paragraph.LineSpacingBeforePoints = 0;
        paragraph.LineSpacingAfterPoints = 0;
        paragraph.LineSpacingRule = WordLineSpacingRule.Auto;
        paragraph.LineSpacing = 240;

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(1, pdf.NumberOfPages);
        var words = pdf.GetPage(1).GetWords().ToList();
        Assert.Equal(text.Split(' '), words.Select(word => word.Text));
        Assert.Equal(3, words.Select(word => Math.Round(word.BoundingBox.Bottom, 2)).Distinct().Count());
        double left = words.Min(word => word.BoundingBox.Left);
        Assert.InRange(words.Max(word => word.BoundingBox.Right) - left, 1D, 140D);
    }
}
