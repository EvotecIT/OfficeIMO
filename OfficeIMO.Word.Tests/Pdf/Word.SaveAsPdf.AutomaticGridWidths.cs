using System.Linq;
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

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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


        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var words = pdf.GetPage(1).GetWords().ToList();
        var left = Assert.Single(words, word => word.Text == "WidthMarker0");
        var right = Assert.Single(words, word => word.Text == "WidthMarker1");
        // The authored 80pt column grows to its actual styled-word minimum;
        // the adjoining 160pt column keeps its authored minimum.
        Assert.InRange(right.BoundingBox.Left - left.BoundingBox.Left, 85D, 88D);
        Assert.Equal(74, words.Count);
    }

    [Theory]
    [InlineData(20, 2, WordTableWidthKind.Length)]
    [InlineData(45, 3, WordTableWidthKind.Length)]
    [InlineData(20, 1, WordTableWidthKind.Auto)]
    [InlineData(45, 3, WordTableWidthKind.Auto)]
    public void SaveAsPdf_AutomaticNoWrapCellUsesPreferredWidthType(int characterCount, int expectedLines, WordTableWidthKind widthType) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(document, 2400);
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.WrapText = false;
        cell.Width = TableWidthFixture(widthType, cell.Width?.ToNativeValue() ?? 5000);
        cell.Paragraphs[0].Text = "WidthMarker0 " + new string('W', characterCount);

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
        if (indent == "left") paragraph.IndentationBefore = global::OfficeIMO.OfficeLength.Points(36D);
        else if (indent == "right") paragraph.IndentationAfter = global::OfficeIMO.OfficeLength.Points(36D);
        else paragraph.IndentationFirstLine = global::OfficeIMO.OfficeLength.Points(36D);

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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

        // This narrow-width fixture pins Base-14 line counts, independently of host fonts.
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false,
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(text, string.Concat(letters.Select(letter => letter.Value)));
        Assert.True(letters.Select(letter => Math.Round(letter.StartBaseLine.Y, 2)).Distinct().Count() >= 4);
        Assert.All(letters, letter => Assert.True(letter.BoundingBox.Right <= 192D));
    }

    [Theory]
    [InlineData(18, false, 80D)]
    [InlineData(24, false, 107.15D)]
    [InlineData(18, true, 65.9D)]
    [InlineData(24, true, 86.75D)]
    public void SaveAsPdf_AutomaticUnequalGridDistributesItsSpanningCellConstraint(
        int characters, bool automaticLeft, double expectedGap) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(document, 1600, 3200);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Span0";
        table.Rows[0].Cells[1].Paragraphs[0].Text = "Span1";
        if (automaticLeft) {
            table.Rows[0].Cells[0].Width = WordTableWidth.Auto;
        }
        WordTableRow row = table.AddRow();
        row.Cells[0].MergeHorizontally(1);
        WordTableCell cell = row.Cells[0];
        cell.Width = WordTableWidth.Auto;
        cell.Paragraphs[0].Text = new string('W', characters);
        cell.Paragraphs[0].FontFamily = "Arial"; cell.Paragraphs[0].FontSize = global::OfficeIMO.OfficeLength.Points(12);
        table.GridColumnWidth = new List<int> { 1600, 3200 }.Select(value => global::OfficeIMO.OfficeLength.Emus(checked((long)value * 635))).ToList();

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var words = pdf.GetPage(1).GetWords().ToList();
        double gap = Assert.Single(words, word => word.Text == "Span1").BoundingBox.Left -
            Assert.Single(words, word => word.Text == "Span0").BoundingBox.Left;
        Assert.InRange(gap, expectedGap - .25D, expectedGap + .25D);
        Assert.Contains(words, word => word.Text == new string('W', characters));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_AutomaticGridKeepsFallbackFontLinesInsideCellBorders(bool minimumSpacing) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = CreateAutomaticWidthControl(document, 2400).Rows[0].Cells[0].Paragraphs[0];
        paragraph.Text = "A " + string.Concat(Enumerable.Repeat("東京", 24));
        paragraph.LineSpacing = minimumSpacing ? WordLineSpacing.AtLeast(OfficeLength.Points(12)) : WordLineSpacing.Multiple(1);
        var options = CreateAutomaticGridFallbackFontOptions();
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, PdfOptions = options,
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
        paragraph.FontFamily = "Arial"; paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(12);
        paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero; paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        paragraph.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Multiple((spacing) / 240D);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, PdfOptions = CreateAutomaticGridFallbackFontOptions(),
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        var letters = pdf.GetPage(1).Letters.Where(letter => letter.Value == "東" || letter.Value == "京").ToList();
        Assert.Equal(6, letters.Count);
        double[] baselines = letters.Select(letter => letter.StartBaseLine.Y).Distinct().OrderByDescending(value => value).ToArray();
        Assert.Equal(3, baselines.Length);
        for (int index = 1; index < baselines.Length; index++)
            Assert.Equal(19.2D * spacing / 240D, baselines[index - 1] - baselines[index], 3);
    }

    private static OfficeIMO.Pdf.PdfCreateOptions CreateAutomaticGridFallbackFontOptions() {
        var options = new OfficeIMO.Pdf.PdfCreateOptions();
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
            cell.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(gridTwips[column]) * 635)));
            cell.MarginLeft = global::OfficeIMO.OfficeLength.Points(5.4D); cell.MarginRight = global::OfficeIMO.OfficeLength.Points(5.4D);
            cell.MarginTop = global::OfficeIMO.OfficeLength.Zero; cell.MarginBottom = global::OfficeIMO.OfficeLength.Zero;
            var paragraph = cell.Paragraphs[0];
            paragraph.FontFamily = "Arial"; paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(12);
            paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero; paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
            paragraph.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Multiple((240) / 240D);
        }
        table.GridColumnWidth = gridTwips.ToList().Select(value => global::OfficeIMO.OfficeLength.Emus(checked((long)value * 635))).ToList();
        return table;
    }
}
