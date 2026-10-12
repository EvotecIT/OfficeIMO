using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0, false, false, false, 118.8D, 163.85D)]
    [InlineData(0, true, false, false, 76.85D, 205.8D)]
    [InlineData(1, false, false, false, 141.35D, 195.75D)]
    [InlineData(1, true, false, false, 86.9D, 195.75D)]
    [InlineData(1, false, true, false, 118.8D, 204.4D)]
    [InlineData(1, true, true, false, 68.85D, 213.8D)]
    [InlineData(0, false, false, true, 61.7D, 143.25D)]
    [InlineData(0, true, false, true, 104D, 178.65D)]
    [InlineData(2, false, false, true, 120D, 172.6D)]
    [InlineData(2, true, false, true, 139.4D, 143.25D)]
    [InlineData(1, false, false, true, 80D, 152.35D)]
    [InlineData(1, true, false, true, 124.25D, 158.4D)]
    public void SaveAsPdf_AutomaticOverlappingMergedCellsDistributeWidthsInSourceOrder(
        int gridShape, bool rightSpanFirst, bool automaticLeft, bool shortSpanFirst,
        double firstWidth, double secondWidth) {
        using WordDocument document = WordDocument.Create();
        int[] preferences = gridShape == 1 ? new[] { 80, 80, 80 }
            : gridShape == 2 ? new[] { 120, 80, 40 } : new[] { 40, 80, 120 };
        WordTable table = CreateAutomaticWidthControl(document, preferences.Select(value => value * 20).ToArray());
        for (int column = 0; column < 3; column++)
            table.Rows[0].Cells[column].Paragraphs[0].Text = "Span" + column;
        if (automaticLeft) {
            table.Rows[0].Cells[0].WidthType = WordTableWidthUnit.Auto;
            table.Rows[0].Cells[0].Width = 0;
        }
        int[] spanStarts = rightSpanFirst ? new[] { 1, 0 } : new[] { 0, 1 };
        for (int row = 0; row < spanStarts.Length; row++)
            AddAutomaticControlSpan(table, preferences, spanStarts[row],
                shortSpanFirst && row == 0 ? 12 : 24, automaticLeft);
        table.GridColumnWidth = preferences.Select(value => value * 20).ToList();
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var page = Assert.Single(pdf.GetPages());
        var words = page.GetWords().ToList();
        var firstRow = words.Where(word => word.Text.StartsWith("Span", StringComparison.Ordinal))
            .GroupBy(word => Math.Round(word.BoundingBox.Bottom, 2)).OrderByDescending(group => group.Key).First().ToList();
        Assert.Equal(3, firstRow.Count);
        Assert.InRange(firstRow[1].BoundingBox.Left - firstRow[0].BoundingBox.Left, firstWidth - .3D, firstWidth + .3D);
        Assert.InRange(firstRow[2].BoundingBox.Left - firstRow[1].BoundingBox.Left, secondWidth - .3D, secondWidth + .3D);
        Assert.Equal(shortSpanFirst ? 1 : 2, words.Count(word => word.Text == new string('W', 24)));
        if (shortSpanFirst) Assert.Single(words, word => word.Text == new string('W', 12));
    }

    [Theory]
    [InlineData(80, 80, 150, 80D)]
    [InlineData(80, 120, 150, 76.55D)]
    [InlineData(120, 80, 150, 85.9D)]
    [InlineData(80, 80, 175, 80D)]
    [InlineData(80, 120, 175, 78.25D)]
    [InlineData(120, 80, 175, 102.8D)]
    public void SaveAsPdf_AutomaticMergedMinimumFitsAConstrainedFrame(
        int leftPreference, int rightPreference, int bodyWidth, double firstWidth) {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].PageSettings.Width = (uint)((bodyWidth + 144) * 20);
        document.Sections[0].Margins.Left = 1440; document.Sections[0].Margins.Right = 1440;
        int[] preferences = { leftPreference, rightPreference };
        WordTable table = CreateAutomaticWidthControl(document, preferences.Select(value => value * 20).ToArray());
        for (int column = 0; column < 2; column++) table.Rows[0].Cells[column].Paragraphs[0].Text = "Span" + column;
        AddAutomaticControlSpan(table, preferences, 0, 12);
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var words = Assert.Single(pdf.GetPages()).GetWords().ToArray();
        var left = Assert.Single(words, word => word.Text == "Span0");
        var right = Assert.Single(words, word => word.Text == "Span1");
        Assert.InRange(right.BoundingBox.Left - left.BoundingBox.Left, firstWidth - .15D, firstWidth + .15D);
        Assert.Single(words, word => word.Text == new string('W', 12));
    }

    [Theory]
    [InlineData(0, 200, 42.6D, 74.25D)]
    [InlineData(0, 220, 45.5D, 80D)]
    [InlineData(1, 200, 92.15D, 64.75D)]
    [InlineData(1, 220, 102.75D, 71.25D)]
    [InlineData(2, 200, 73.55D, 65.6D)]
    [InlineData(2, 220, 79.45D, 72.35D)]
    public void SaveAsPdf_AutomaticThreeColumnSpanRespectsPreferencesAndPadding(
        int gridShape, int bodyWidth, double firstWidth, double secondWidth) {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].PageSettings.Width = (uint)((bodyWidth + 144) * 20);
        document.Sections[0].Margins.Left = 1440; document.Sections[0].Margins.Right = 1440;
        int[] preferences = gridShape == 0 ? new[] { 40, 80, 120 }
            : gridShape == 1 ? new[] { 120, 80, 40 } : new[] { 80, 80, 80 };
        WordTable table = CreateAutomaticWidthControl(document, preferences.Select(value => value * 20).ToArray());
        for (int column = 0; column < 3; column++) table.Rows[0].Cells[column].Paragraphs[0].Text = "Span" + column;
        AddAutomaticControlSpan(table, preferences, 0, 18, columnSpan: 3);
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        var letters = Assert.Single(pdf.GetPages()).Letters;
        var starts = letters.Where(letter => letter.Value == "S").OrderBy(letter => letter.StartBaseLine.X).ToArray();
        Assert.Equal(3, starts.Length);
        Assert.InRange(starts[1].StartBaseLine.X - starts[0].StartBaseLine.X, firstWidth - .15D, firstWidth + .15D);
        Assert.InRange(starts[2].StartBaseLine.X - starts[1].StartBaseLine.X, secondWidth - .15D, secondWidth + .15D);
        Assert.Equal("Span0Span1Span2" + new string('W', 18),
            string.Concat(letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).Select(letter => letter.Value)));
    }

    private static void AddAutomaticControlSpan(WordTable table, int[] preferences, int startColumn,
        int characterCount, bool automaticLeft = false, int columnSpan = 2) {
        WordTableRow row = table.AddRow();
        for (int column = 0; column < preferences.Length; column++) {
            WordTableCell cell = row.Cells[column];
            cell.WidthType = automaticLeft && column == 0 ? WordTableWidthUnit.Auto : WordTableWidthUnit.Dxa;
            cell.Width = cell.WidthType == WordTableWidthUnit.Auto ? 0 : preferences[column] * 20;
            cell.MarginLeftWidth = 108; cell.MarginRightWidth = 108;
            cell.MarginTopWidth = 0; cell.MarginBottomWidth = 0;
            var paragraph = cell.Paragraphs[0];
            paragraph.Text = "Span" + column;
            paragraph.FontFamily = "Arial"; paragraph.FontSize = 12;
            paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
            paragraph.LineSpacingRule = WordLineSpacingRule.Auto; paragraph.LineSpacing = 240;
        }
        WordTableCell merged = row.Cells[startColumn];
        merged.MergeHorizontally(columnSpan - 1);
        merged.WidthType = WordTableWidthUnit.Auto; merged.Width = 0;
        merged.Paragraphs[0].Text = new string('W', characterCount);
    }

    [Theory]
    [InlineData(2000, 80, 0, 0, 62.05D)]
    [InlineData(6000, 80, 0, 0, 62.05D)]
    [InlineData(2000, 80, 0, 1, 70.95D)]
    [InlineData(6000, 80, 120, 2, 49.462D)]
    [InlineData(2000, 500, 0, 0, 62.036D)]
    [InlineData(6000, 500, 0, 1, 112.466D)]
    [InlineData(2000, 500, 120, 0, 74.036D)]
    [InlineData(6000, 500, 120, 1, 124.466D)]
    public void SaveAsPdf_AutomaticTableFrameUsesCellPreferencesAndTheAvailableBody(
        int cachedGridTwips, int bodyWidth, int spacingTwips, int automaticColumn, double expectedAdvance) {
        using WordDocument document = WordDocument.Create();
        var section = document.Sections[0];
        section.PageSettings.Width = (uint)((bodyWidth + 144) * 20);
        section.Margins.Left = 1440; section.Margins.Right = 1440;
        WordTable table = document.AddTable(1, 2, WordTableStyle.TableGrid);
        table.WidthType = WordTableWidthUnit.Auto; table.Width = 0;
        table._tableProperties.TableLayout?.Remove();
        table._tableProperties.TableBorders = new W.TableBorders(
            new W.TopBorder { Val = W.BorderValues.Single, Size = 8U },
            new W.LeftBorder { Val = W.BorderValues.Single, Size = 8U },
            new W.BottomBorder { Val = W.BorderValues.Single, Size = 8U },
            new W.RightBorder { Val = W.BorderValues.Single, Size = 8U },
            new W.InsideHorizontalBorder { Val = W.BorderValues.Single, Size = 8U },
            new W.InsideVerticalBorder { Val = W.BorderValues.Single, Size = 8U });
        table._tableProperties.TableCellSpacing = new W.TableCellSpacing {
            Type = W.TableWidthUnitValues.Dxa, Width = spacingTwips.ToString(System.Globalization.CultureInfo.InvariantCulture)
        };
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.WidthType = automaticColumn == column ? WordTableWidthUnit.Auto : WordTableWidthUnit.Dxa;
            cell.Width = automaticColumn == column ? 0 : automaticColumn == 2 ? (column == 0 ? 1200 : 3300) : 2250;
            cell.MarginLeftWidth = 0; cell.MarginRightWidth = 0;
            cell.MarginTopWidth = 0; cell.MarginBottomWidth = 0;
            var paragraph = cell.Paragraphs[0];
            paragraph.Text = column == 0 ? "ABCDEFGHIJ" : "Z";
            paragraph.FontFamily = "Courier New"; paragraph.FontSize = 10;
            paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
        }
        table.GridColumnWidth = new List<int> { cachedGridTwips / 2, cachedGridTwips / 2 };
        using var pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PdfOptions = new PdfOptions { DefaultFont = PdfStandardFont.Courier }
        }));
        var page = Assert.Single(pdf.GetPages());
        Assert.Equal("ABCDEFGHIJZ", string.Concat(page.Letters.Select(letter => letter.Value)));
        var first = Assert.Single(page.Letters, letter => letter.Value == "A");
        var second = Assert.Single(page.Letters, letter => letter.Value == "Z");
        Assert.InRange(second.StartBaseLine.X - first.StartBaseLine.X, expectedAdvance - .15D, expectedAdvance + .15D);
        Assert.All(page.Letters, letter => Assert.InRange(letter.BoundingBox.Right, 72D, 72D + bodyWidth));
    }
}
