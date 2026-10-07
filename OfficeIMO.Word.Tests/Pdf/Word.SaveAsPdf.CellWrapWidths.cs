using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordTableLayoutMode.Fixed, WordTableWidthUnit.Dxa, 3)]
    [InlineData(WordTableLayoutMode.Fixed, WordTableWidthUnit.Auto, 3)]
    [InlineData(WordTableLayoutMode.Fixed, WordTableWidthUnit.Pct, 3)]
    [InlineData(WordTableLayoutMode.AutoFit, WordTableWidthUnit.Dxa, 3)]
    [InlineData(WordTableLayoutMode.AutoFit, WordTableWidthUnit.Auto, 1)]
    [InlineData(WordTableLayoutMode.AutoFit, WordTableWidthUnit.Pct, 1)]
    public void SaveAsPdf_CellNoWrapUsesTableLayoutAndPreferredWidthType(
        WordTableLayoutMode layout, WordTableWidthUnit widthType, int expectedLines) {
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
        Assert.Equal(expectedLines, words.Select(word => Math.Round(word.BoundingBox.Bottom, 2)).Distinct().Count());
        double left = words.Min(word => word.BoundingBox.Left);
        Assert.InRange(words.Max(word => word.BoundingBox.Right) - left, 1D, expectedLines == 1 ? 468D : 140D);
        var frame = Assert.Single(pdf.GetPage(1).Paths.Where(path => path.IsStroked)
            .Select(path => path.GetBoundingRectangle()).Where(bounds => bounds.HasValue)
            .Select(bounds => bounds!.Value));
        // Extractors can return text outside a PDF clipping rectangle. Every word
        // must also fit inside the visible cell, including its authored padding.
        Assert.True(words.Max(word => word.BoundingBox.Right) <= frame.Right - 5.4D + 0.01D);
    }

    [Theory]
    [InlineData(WordTableWidthUnit.Dxa, 6000, WordTableWidthUnit.Dxa, true, 3, 150D)]
    [InlineData(WordTableWidthUnit.Dxa, 6000, WordTableWidthUnit.Auto, true, 2, 260.5D)]
    [InlineData(WordTableWidthUnit.Dxa, 6000, WordTableWidthUnit.Auto, false, 1, 301D)]
    [InlineData(WordTableWidthUnit.Pct, 5000, WordTableWidthUnit.Dxa, true, 2, 234D)]
    [InlineData(WordTableWidthUnit.Pct, 5000, WordTableWidthUnit.Auto, true, 1, 413.4D)]
    [InlineData(WordTableWidthUnit.Pct, 5000, WordTableWidthUnit.Auto, false, 1, 413.4D)]
    [InlineData(WordTableWidthUnit.Dxa, 3000, WordTableWidthUnit.Dxa, true, 7, 75D)]
    [InlineData(WordTableWidthUnit.Dxa, 3000, WordTableWidthUnit.Auto, true, 4, 110.5D)]
    [InlineData(WordTableWidthUnit.Dxa, 3000, WordTableWidthUnit.Auto, false, 1, 301D)]
    public void SaveAsPdf_AutomaticTableSharesPreferredWidthAccordingToCellPreferences(
        WordTableWidthUnit tableWidthType, int tableWidth, WordTableWidthUnit cellWidthType,
        bool wrap, int expectedLines, double expectedGap) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 2);
        table.ConditionalFormattingFirstRow = false;
        table.WidthType = tableWidthType;
        table.Width = tableWidth;
        table.LayoutMode = WordTableLayoutMode.AutoFit;
        int gridWidth = tableWidthType == WordTableWidthUnit.Pct ? 4680 : tableWidth / 2;
        table.GridColumnWidth = new List<int> { gridWidth, gridWidth };
        const string text = "Start Alpha Beta Gamma Delta Epsilon Zeta Eta Theta";
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.Width = cellWidthType == WordTableWidthUnit.Dxa ? gridWidth : 0;
            cell.WidthType = cellWidthType;
            cell.WrapText = wrap;
            cell.MarginLeftWidth = 108;
            cell.MarginRightWidth = 108;
            cell.MarginTopWidth = 0;
            cell.MarginBottomWidth = 0;
            WordParagraph paragraph = cell.Paragraphs[0];
            paragraph.Text = column == 0 ? text : "Short";
            paragraph.FontFamily = "Arial";
            paragraph.FontSize = 12;
            paragraph.LineSpacingBeforePoints = 0;
            paragraph.LineSpacingAfterPoints = 0;
            paragraph.LineSpacingRule = WordLineSpacingRule.Auto;
            paragraph.LineSpacing = 240;
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(1, pdf.NumberOfPages);
        var words = pdf.GetPage(1).GetWords().ToList();
        var shortWord = Assert.Single(words, word => word.Text == "Short");
        var leftWords = words.Where(word => word != shortWord).ToList();
        Assert.Equal(text.Split(' '), leftWords.Select(word => word.Text));
        Assert.Equal(expectedLines, leftWords.Select(word => Math.Round(word.BoundingBox.Bottom, 2)).Distinct().Count());
        double gap = shortWord.BoundingBox.Left - leftWords.Min(word => word.BoundingBox.Left);
        // Independent Word exports allow small font-metric rounding differences.
        Assert.InRange(gap, expectedGap - 1D, expectedGap + 1D);
        Assert.True(leftWords.Max(word => word.BoundingBox.Right) <= shortWord.BoundingBox.Left - 10.8D + .01D);
    }
}
