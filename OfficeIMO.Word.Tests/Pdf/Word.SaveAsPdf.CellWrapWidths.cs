using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordTableLayoutMode.Fixed, WordTableWidthKind.Length, 3)]
    [InlineData(WordTableLayoutMode.Fixed, WordTableWidthKind.Auto, 3)]
    [InlineData(WordTableLayoutMode.Fixed, WordTableWidthKind.Percent, 3)]
    [InlineData(WordTableLayoutMode.AutoFit, WordTableWidthKind.Length, 3)]
    [InlineData(WordTableLayoutMode.AutoFit, WordTableWidthKind.Auto, 1)]
    [InlineData(WordTableLayoutMode.AutoFit, WordTableWidthKind.Percent, 1)]
    public void SaveAsPdf_CellNoWrapUsesTableLayoutAndPreferredWidthType(
        WordTableLayoutMode layout, WordTableWidthKind widthType, int expectedLines) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        table.ConditionalFormattingFirstRow = false;
        table.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(3000) * 635)));
        table.LayoutMode = layout;
        table.GridColumnWidth = new List<int> { 3000 }.Select(value => global::OfficeIMO.OfficeLength.Emus(checked((long)value * 635))).ToList();
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.Width = TableWidthFixture(widthType, widthType == WordTableWidthKind.Percent ? 5000 : 3000);
        cell.WrapText = false;
        cell.MarginLeft = global::OfficeIMO.OfficeLength.Points(5.4D);
        cell.MarginRight = global::OfficeIMO.OfficeLength.Points(5.4D);
        cell.MarginTop = global::OfficeIMO.OfficeLength.Zero;
        cell.MarginBottom = global::OfficeIMO.OfficeLength.Zero;
        WordParagraph paragraph = cell.Paragraphs[0];
        const string text = "Start Alpha Beta Gamma Delta Epsilon Zeta Eta Theta";
        paragraph.Text = text;
        paragraph.FontFamily = "Arial";
        paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(12);
        paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;
        paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
        paragraph.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Multiple((240) / 240D);

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
    [InlineData(WordTableWidthKind.Length, 6000, WordTableWidthKind.Length, true, 3, 150D)]
    [InlineData(WordTableWidthKind.Length, 6000, WordTableWidthKind.Auto, true, 2, 260.5D)]
    [InlineData(WordTableWidthKind.Length, 6000, WordTableWidthKind.Auto, false, 1, 301D)]
    [InlineData(WordTableWidthKind.Percent, 5000, WordTableWidthKind.Length, true, 2, 234D)]
    [InlineData(WordTableWidthKind.Percent, 5000, WordTableWidthKind.Auto, true, 1, 413.4D)]
    [InlineData(WordTableWidthKind.Percent, 5000, WordTableWidthKind.Auto, false, 1, 413.4D)]
    [InlineData(WordTableWidthKind.Length, 3000, WordTableWidthKind.Length, true, 7, 75D)]
    [InlineData(WordTableWidthKind.Length, 3000, WordTableWidthKind.Auto, true, 4, 110.5D)]
    [InlineData(WordTableWidthKind.Length, 3000, WordTableWidthKind.Auto, false, 1, 301D)]
    public void SaveAsPdf_AutomaticTableSharesPreferredWidthAccordingToCellPreferences(
        WordTableWidthKind tableWidthType, int tableWidth, WordTableWidthKind cellWidthType,
        bool wrap, int expectedLines, double expectedGap) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 2);
        table.ConditionalFormattingFirstRow = false;
        table.Width = TableWidthFixture(tableWidthType, tableWidth);
        table.LayoutMode = WordTableLayoutMode.AutoFit;
        int gridWidth = tableWidthType == WordTableWidthKind.Percent ? 4680 : tableWidth / 2;
        table.GridColumnWidth = new List<int> { gridWidth, gridWidth }.Select(value => global::OfficeIMO.OfficeLength.Emus(checked((long)value * 635))).ToList();
        const string text = "Start Alpha Beta Gamma Delta Epsilon Zeta Eta Theta";
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.Width = TableWidthFixture(cellWidthType, cellWidthType == WordTableWidthKind.Length ? gridWidth : 0);
            cell.WrapText = wrap;
            cell.MarginLeft = global::OfficeIMO.OfficeLength.Points(5.4D);
            cell.MarginRight = global::OfficeIMO.OfficeLength.Points(5.4D);
            cell.MarginTop = global::OfficeIMO.OfficeLength.Zero;
            cell.MarginBottom = global::OfficeIMO.OfficeLength.Zero;
            WordParagraph paragraph = cell.Paragraphs[0];
            paragraph.Text = column == 0 ? text : "Short";
            paragraph.FontFamily = "Arial";
            paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(12);
            paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;
            paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
            paragraph.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Multiple((240) / 240D);
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
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
