using System.IO;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfIdentifierSeparatorTests {
    [Fact]
    public void SeparatorBreaks_Hostname_PreservesWholeLabels() {
        const string source = "Mx12.LongLabel-mail7.example.test";
        const string expected = "Mx12.\u200BLongLabel-\u200Bmail7.\u200Bexample.\u200Btest";

        string display = WordPaperText.WithIdentifierBreaks(source, '.', '-');

        Assert.Equal(expected, display);
        Assert.Equal(source, WordPaperText.RemoveOptionalBreaks(display));
        Assert.Equal(display, WordPaperText.WithIdentifierBreaks(display, '.', '-'));
        Assert.Equal("LongLabel123", WordPaperText.WithIdentifierBreaks("LongLabel123", '.', '-'));
    }

    [Fact]
    public void SeparatorBreaks_WrappedWordAndPdf_PreserveExactHostname() {
        const string source = "Mx12.LongLabel-mail7.example.test";
        const string expected = "Mx12.\u200BLongLabel-\u200Bmail7.\u200Bexample.\u200Btest";
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 2);
        new WordPaperStyle().SetCellText(table.Rows[0].Cells[0], WordPaperText.WithIdentifierBreaks(source, '.', '-'));
        table.Rows[0].Cells[0].Paragraphs[0].FontFamily = "Consolas";
        table.SetColumnWidthsPercentage(18, 82);
        using MemoryStream word = new();

        document.Save(word);
        word.Position = 0;
        using WordDocument reopened = WordDocument.Load(word);
        string stored = reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text;

        Assert.Equal(expected, stored);
        Assert.Equal(source, WordPaperText.RemoveOptionalBreaks(stored));
        byte[] pdf = reopened.ToPdfDocument(new WordToPdfOptions { IncludePageNumbers = false }).ToBytes();
        string extracted = PdfReadDocument.Open(pdf).ExtractText().Trim();
        Assert.Equal(source, extracted);
        Assert.DoesNotContain('\u200B', extracted);
    }
}
