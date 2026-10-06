using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(2400, 13)]
    [InlineData(4800, 6)]
    public void SaveAsPdf_AutomaticGridRetainsAuthoredWidthForBreakableText(int gridTwips, int expectedLines) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(document, gridTwips);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "WidthMarker0 " +
            string.Join(" ", Enumerable.Repeat("wrapped cell text", 12));

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(1, pdf.NumberOfPages);
        var words = pdf.GetPage(1).GetWords().ToList();
        Assert.Single(words, word => word.Text == "WidthMarker0");
        Assert.Equal(37, words.Count);
        Assert.Equal(expectedLines, words.Select(word => Math.Round(word.BoundingBox.Bottom, 2)).Distinct().Count());
        double left = words.Min(word => word.BoundingBox.Left);
        Assert.InRange(words.Max(word => word.BoundingBox.Right) - left, 1D, gridTwips / 20D - 10D);
    }

    [Fact]
    public void SaveAsPdf_AutomaticUnequalGridGrowsOnlyTheColumnWhoseWordDoesNotFit() {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(document, 1600, 3200);
        for (int column = 0; column < 2; column++)
            table.Rows[0].Cells[column].Paragraphs[0].Text = $"WidthMarker{column} " +
                string.Join(" ", Enumerable.Repeat("wrapped cell text", 12));


        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var words = pdf.GetPage(1).GetWords().ToList();
        var left = Assert.Single(words, word => word.Text == "WidthMarker0");
        var right = Assert.Single(words, word => word.Text == "WidthMarker1");
        // The authored 80pt column grows to its actual styled-word minimum;
        // the adjoining 160pt column keeps its authored minimum.
        Assert.InRange(right.BoundingBox.Left - left.BoundingBox.Left, 85D, 88D);
        Assert.Equal(74, words.Count);
    }

    [Theory]
    [InlineData(20, 1)]
    [InlineData(45, 3)]
    public void SaveAsPdf_AutomaticNoWrapCellWrapsOnlyWhenItsTextExceedsThePage(int characterCount, int expectedLines) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(document, 2400);
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.WrapText = false;
        cell.Paragraphs[0].Text = "WidthMarker0 " + new string('W', characterCount);

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(1, pdf.NumberOfPages);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(characterCount, letters.Count(letter => letter.Value == "W") - 1);
        Assert.Equal(expectedLines, letters.Select(letter => Math.Round(letter.StartBaseLine.Y, 2)).Distinct().Count());
        Assert.All(letters, letter => Assert.True(letter.BoundingBox.Right <= 540D,
            $"No-wrap text exceeds the page's right content edge: {letter.BoundingBox.Right}"));
    }

    private static WordTable CreateAutomaticWidthControl(WordDocument document, params int[] gridTwips) {
        WordTable table = document.AddTable(1, gridTwips.Length);
        table.ConditionalFormattingFirstRow = false;
        table.AutoFitToContents();
        for (int column = 0; column < gridTwips.Length; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.Width = gridTwips[column]; cell.WidthType = WordTableWidthUnit.Dxa;
            cell.MarginLeftWidth = 108; cell.MarginRightWidth = 108;
            cell.MarginTopWidth = 0; cell.MarginBottomWidth = 0;
            var paragraph = cell.Paragraphs[0];
            paragraph.FontFamily = "Arial"; paragraph.FontSize = 12;
            paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
            paragraph.LineSpacingRule = WordLineSpacingRule.Auto; paragraph.LineSpacing = 240;
        }
        table.GridColumnWidth = gridTwips.ToList();
        return table;
    }
}
