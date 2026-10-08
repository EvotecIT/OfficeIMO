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
    [InlineData(20, 2, WordTableWidthUnit.Dxa)]
    [InlineData(45, 3, WordTableWidthUnit.Dxa)]
    [InlineData(20, 1, WordTableWidthUnit.Auto)]
    [InlineData(45, 3, WordTableWidthUnit.Auto)]
    public void SaveAsPdf_AutomaticNoWrapCellUsesPreferredWidthType(int characterCount, int expectedLines, WordTableWidthUnit widthType) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(document, 2400);
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.WrapText = false;
        cell.WidthType = widthType;
        cell.Paragraphs[0].Text = "WidthMarker0 " + new string('W', characterCount);

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(1, pdf.NumberOfPages);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(characterCount, letters.Count(letter => letter.Value == "W") - 1);
        Assert.Equal(expectedLines, letters.Select(letter => Math.Round(letter.StartBaseLine.Y, 2)).Distinct().Count());
        Assert.All(letters, letter => Assert.True(letter.BoundingBox.Right <= 540D,
            $"No-wrap text exceeds the page's right content edge: {letter.BoundingBox.Right}"));
    }

    [Theory]
    [InlineData("left")]
    [InlineData("right")]
    [InlineData("first")]
    public void SaveAsPdf_AutomaticNoWrapCellIncludesParagraphIndentsInItsFrame(string indent) {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = CreateAutomaticWidthControl(document, 2400).Rows[0].Cells[0];
        cell.WrapText = false;
        WordParagraph paragraph = cell.Paragraphs[0];
        paragraph.Text = new string('W', 40);
        if (indent == "left") paragraph.IndentationBeforePoints = 36;
        else if (indent == "right") paragraph.IndentationAfterPoints = 36;
        else paragraph.IndentationFirstLinePoints = 36;

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(40, letters.Count(letter => letter.Value == "W"));
        Assert.Equal(2, letters.Select(letter => Math.Round(letter.StartBaseLine.Y, 2)).Distinct().Count());
        Assert.All(letters, letter => Assert.True(letter.BoundingBox.Right <= 540D));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_AutomaticGridUsesHyphenBreaksAcrossStyledRuns(bool styledRuns) {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = CreateAutomaticWidthControl(document, 2400).Rows[0].Cells[0];
        string text = string.Concat(Enumerable.Repeat("part-", 12)) + "end";
        cell.Paragraphs[0].Text = styledRuns ? text.Substring(0, 27) : text;
        if (styledRuns) cell.Paragraphs[0].AddText(text.Substring(27)).Bold = true;

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(text, string.Concat(letters.Select(letter => letter.Value)));
        Assert.True(letters.Select(letter => Math.Round(letter.StartBaseLine.Y, 2)).Distinct().Count() >= 4);
        Assert.All(letters, letter => Assert.True(letter.BoundingBox.Right <= 192D));
    }

    [Theory]
    [InlineData(18, 80D)]
    [InlineData(24, 94D)]
    public void SaveAsPdf_AutomaticUnequalGridGrowsOnlyByItsSpanningCellDeficit(int characters, double expectedGap) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(document, 1600, 3200);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Span0";
        table.Rows[0].Cells[1].Paragraphs[0].Text = "Span1";
        WordTableRow row = table.AddRow();
        row.Cells[0].MergeHorizontally(1);
        WordTableCell cell = row.Cells[0];
        cell.Width = 0; cell.WidthType = WordTableWidthUnit.Auto;
        cell.Paragraphs[0].Text = new string('W', characters);
        cell.Paragraphs[0].FontFamily = "Arial"; cell.Paragraphs[0].FontSize = 12;
        table.GridColumnWidth = new List<int> { 1600, 3200 };

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var words = pdf.GetPage(1).GetWords().ToList();
        double gap = Assert.Single(words, word => word.Text == "Span1").BoundingBox.Left -
            Assert.Single(words, word => word.Text == "Span0").BoundingBox.Left;
        Assert.InRange(gap, expectedGap, expectedGap + 1D);
        Assert.Contains(words, word => word.Text == new string('W', characters));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_AutomaticGridKeepsFallbackFontLinesInsideCellBorders(bool minimumSpacing) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = CreateAutomaticWidthControl(document, 2400).Rows[0].Cells[0].Paragraphs[0];
        paragraph.Text = "A " + string.Concat(Enumerable.Repeat("東京", 24));
        paragraph.LineSpacingRule = minimumSpacing ? WordLineSpacingRule.AtLeast : WordLineSpacingRule.Auto;
        paragraph.LineSpacing = 240;
        var options = CreateAutomaticGridFallbackFontOptions();
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, PdfOptions = options,
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        }));
        var page = pdf.GetPage(1);
        var letters = page.Letters.Where(letter => letter.Value == "東" || letter.Value == "京").ToList();
        Assert.Equal(48, letters.Count);
        double[] baselines = letters.Select(letter => letter.StartBaseLine.Y).Distinct().OrderByDescending(value => value).ToArray();
        Assert.True(baselines.Length >= 3);
        for (int index = 1; index < baselines.Length; index++)
            Assert.Equal(19.2D, baselines[index - 1] - baselines[index], 3);
        double bottom = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(rectangle => rectangle.HasValue).Min(rectangle => rectangle!.Value.Bottom);
        Assert.All(letters, letter => Assert.True(letter.BoundingBox.Bottom >= bottom,
            $"Fallback glyph below table clip: {letter.BoundingBox.Bottom} < {bottom}"));
    }

    [Theory]
    [InlineData(false, 360)]
    [InlineData(true, 360)]
    [InlineData(false, 120)]
    [InlineData(true, 120)]
    public void SaveAsPdf_TableListSpacerPreservesFallbackFontAutomaticMultiple(bool numbered, int spacing) {
        using WordDocument document = WordDocument.Create();
        WordTableCell cell = CreateAutomaticWidthControl(document, 2400).Rows[0].Cells[0];
        WordParagraph paragraph = numbered ? cell.AddList(WordListStyle.Numbered).AddItem(string.Empty) : cell.Paragraphs[0];
        paragraph.Text = "東京\n東京\n東京";
        paragraph.FontFamily = "Arial"; paragraph.FontSize = 12;
        paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
        paragraph.LineSpacingRule = WordLineSpacingRule.Auto; paragraph.LineSpacing = spacing;
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, PdfOptions = CreateAutomaticGridFallbackFontOptions(),
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        }));
        var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value == "東" || letter.Value == "京").ToList();
        Assert.Equal(6, letters.Count);
        double[] baselines = letters.Select(letter => letter.StartBaseLine.Y).Distinct().OrderByDescending(value => value).ToArray();
        Assert.Equal(3, baselines.Length);
        for (int index = 1; index < baselines.Length; index++)
            Assert.Equal(19.2D * spacing / 240D, baselines[index - 1] - baselines[index], 3);
    }

    private static OfficeIMO.Pdf.PdfOptions CreateAutomaticGridFallbackFontOptions() {
        var options = new OfficeIMO.Pdf.PdfOptions();
        options.RegisterNamedFontFamily(new OfficeIMO.Pdf.PdfEmbeddedFontFamily("Arial", CreateBaselineMetricFont()));
        options.RegisterEmbeddedFontFallbacks(new OfficeIMO.Pdf.PdfEmbeddedFontFallbackSet(new[] {
            new OfficeIMO.Pdf.PdfEmbeddedFontFallbackCandidate("Tall CJK",
                OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(1200, -300, 100, 300, ' ', '東', '京'))
        }));
        return options;
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
