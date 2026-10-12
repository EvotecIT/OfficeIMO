using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("together")]
    [InlineData("next")]
    [InlineData("widow")]
    public void SaveAsPdf_MergedCellKeepsItsSourceParagraphPagination(string rule) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(3, 2);
        table.ConditionalFormattingFirstRow = false;
        int firstCount = rule == "next" ? 3 : 4;
        WordTableCell anchor = table.Rows[0].Cells[0];
        WordParagraph first = anchor.Paragraphs[0];
        first.Text = "First1";
        for (int n = 2; n <= firstCount; n++) { first.AddBreak(); first.AddText($"First{n}"); }
        WordParagraph second = anchor.AddParagraph("Second1");
        for (int n = 2; n <= 8 - firstCount; n++) { second.AddBreak(); second.AddText($"Second{n}"); }
        for (int row = 0; row < 3; row++) {
            table.Rows[row].Cells[1].Paragraphs[0].Text = $"Neighbour{row}";
            foreach (WordTableCell cell in table.Rows[row].Cells) {
                cell.MarginTopWidth = cell.MarginBottomWidth = 0;
                foreach (WordParagraph paragraph in cell.Paragraphs) {
                    paragraph.FontFamily = "Arial"; paragraph.FontSize = 12;
                    paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0;
                    paragraph.LineSpacingRule = WordLineSpacingRule.Exact; paragraph.LineSpacing = 400;
                    paragraph.KeepLinesTogetherOverride = false; paragraph.KeepWithNextOverride = false;
                    paragraph.AvoidWidowAndOrphanOverride = false;
                }
            }
        }
        first.KeepLinesTogetherOverride = rule == "together";
        first.KeepWithNextOverride = rule == "next";
        first.AvoidWidowAndOrphanOverride = rule == "widow";
        table.MergeCells(0, 0, 3, 1, copyParagraphs: true);
        Assert.Empty(document.ValidateDocument());
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, PageSize = new OfficeIMO.Pdf.PageSize(300, 140),
            Margins = OfficeIMO.Pdf.PageMargins.Uniform(20)
        }));
        var words = pdf.GetPages().SelectMany(page => page.GetWords().Select(word => (page.Number, word.Text))).ToArray();
        int PageOf(string marker) => Assert.Single(words, word => word.Text == marker).Number;
        if (rule == "next") Assert.Equal(PageOf("First3"), PageOf("Second1"));
        else if (rule == "together") Assert.Equal(PageOf("First1"), PageOf("First4"));
        else Assert.All(Enumerable.Range(1, firstCount).Select(n => PageOf($"First{n}")).GroupBy(page => page), group => Assert.True(group.Count() >= 2));
    }

    [Theory]
    [InlineData(2)]
    [InlineData(4)]
    [InlineData(6)]
    public void SaveAsPdf_VerticallyMergedCellTextContinuesWithinPageMargins(int paragraphsPerRow) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(3, 2);
        table.ConditionalFormattingFirstRow = false;
        table.StyleDetails!.SetBordersForAllSides(WordBorderStyle.Single, 4U, OfficeIMO.Drawing.OfficeColor.Black);
        for (int row = 0; row < 3; row++) {
            for (int column = 0; column < 2; column++) {
                WordTableCell cell = table.Rows[row].Cells[column];
                cell.MarginTopWidth = cell.MarginBottomWidth = 0;
                cell.Paragraphs[0].Text = column == 0 ? $"SpanToken{row * paragraphsPerRow + 1}" : $"Neighbour{row}";
                if (column == 0)
                    for (int paragraph = 1; paragraph < paragraphsPerRow; paragraph++)
                        cell.AddParagraph($"SpanToken{row * paragraphsPerRow + paragraph + 1}");
                foreach (WordParagraph paragraph in cell.Paragraphs) {
                    paragraph.FontFamily = "Arial";
                    paragraph.FontSize = 12;
                    paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0;
                    paragraph.LineSpacingRule = WordLineSpacingRule.Auto;
                    paragraph.LineSpacing = 240;
                }
            }
        }
        table.MergeCells(0, 0, 3, 1, copyParagraphs: true);
        Assert.Empty(document.ValidateDocument());
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, PageSize = new OfficeIMO.Pdf.PageSize(300, 120),
            Margins = OfficeIMO.Pdf.PageMargins.Uniform(20)
        }));
        Assert.True(pdf.NumberOfPages > 1);
        var words = pdf.GetPages().SelectMany(page => page.GetWords().Select(word => (page.Number, Word: word))).ToArray();
        for (int token = 1; token <= paragraphsPerRow * 3; token++) {
            Assert.True(words.Count(item => item.Word.Text == $"SpanToken{token}") == 1,
                $"SpanToken{token} must appear once: " + string.Join("; ", words.Select(item => $"{item.Word.Text}@{item.Number}:{item.Word.BoundingBox.Bottom:F2}")));
            var occurrence = Assert.Single(words, item => item.Word.Text == $"SpanToken{token}");
            // Glyph descent can extend slightly outside the nominal text line box.
            Assert.True(occurrence.Word.BoundingBox.Bottom >= 19D,
                $"SpanToken{token} extends below the body on page {occurrence.Number}: {occurrence.Word.BoundingBox.Bottom}.");
        }
        for (int row = 0; row < 3; row++)
            Assert.Single(words, item => item.Word.Text == $"Neighbour{row}");
        Assert.True(Assert.Single(words, item => item.Word.Text == $"SpanToken{paragraphsPerRow * 3}").Number > 1);
    }
}
