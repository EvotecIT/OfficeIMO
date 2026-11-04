using System;
using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public sealed class WordPdfContentsAlignmentTests {
    [Fact]
    public void Contents_IndentedLevelsAlignFinalPageNumbersAtTheSameRightEdge() {
        using WordDocument document = WordDocument.Create();
        document.AddTableOfContent(WordTableOfContentsStyle.Template1, 1, 3);
        document.AddPageBreak();
        document.AddParagraph("Root entry").Style = WordParagraphStyles.Heading1;
        document.AddParagraph("Child entry").Style = WordParagraphStyles.Heading2;
        document.AddParagraph("Leaf entry").Style = WordParagraphStyles.Heading3;
        byte[] bytes = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        var lines = pdf.GetPage(1).Letters.GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 2))
            .Select(group => new { Text = string.Concat(group.Select(letter => letter.Value)), Letters = group.ToArray() }).ToArray();
        var entries = new[] { "Root", "Child", "Leaf" }
            .Select(title => lines.Single(line => line.Text.Contains(title))).ToArray();
        double[] left = entries.Select(entry => entry.Letters[0].StartBaseLine.X).ToArray();
        double[] right = entries.Select(entry => Assert.Single(entry.Letters, letter => letter.Value == "2").GlyphRectangle.Right).ToArray();
        Assert.True(left[0] < left[1] && left[1] < left[2]);
        Assert.All(right, value => Assert.InRange(Math.Abs(value - right[0]), 0D, 0.1D));
    }
}
