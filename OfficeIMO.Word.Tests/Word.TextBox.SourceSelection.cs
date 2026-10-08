using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("drawing", "body")]
    [InlineData("vml", "body")]
    [InlineData("mixed", "body")]
    [InlineData("drawing", "header")]
    [InlineData("vml", "header")]
    [InlineData("mixed", "header")]
    [InlineData("drawing", "footer")]
    [InlineData("vml", "footer")]
    [InlineData("mixed", "footer")]
    public void TextBoxWrappersSelectTheirOwnSourceInSharedRun(string kind, string story) {
        using WordDocument document = WordDocument.Create();
        WordParagraph host = story == "header" ? document.HeaderDefaultOrCreate.AddParagraph("Host")
            : story == "footer" ? document.FooterDefaultOrCreate.AddParagraph("Host") : document.AddParagraph("Host");
        WordTextBox first = kind == "vml" ? host.AddTextBoxVml("First") : host.AddTextBox("First", WordImageTextWrapping.Square);
        WordTextBox second = kind == "drawing" ? host.AddTextBox("Second", WordImageTextWrapping.Square) : host.AddTextBoxVml("Second");
        Assert.Equal("First", first.Paragraphs.Single().Text.TrimEnd('\n'));
        Assert.Equal("Second", second.Paragraphs.Single().Text.TrimEnd('\n'));
        first.Paragraphs.Single().Text = "First edited";
        second.Paragraphs.Single().Text = "Second edited";
        first.Description = "First description";
        second.Description = "Second description";
        Assert.Equal("First edited", first.Paragraphs.Single().Text.TrimEnd('\n'));
        Assert.Equal("Second edited", second.Paragraphs.Single().Text.TrimEnd('\n'));
        var parent = Assert.IsType<WordTextBox>(second.Paragraphs.Single().Parent);
        Assert.Equal("Second edited", parent.Paragraphs.Single().Text.TrimEnd('\n'));
        using WordDocument reopened = WordDocument.Load(new MemoryStream(document.ToBytes()));
        var boxes = story == "header" ? reopened.Header!.Default!.TextBoxes
            : story == "footer" ? reopened.Footer!.Default!.TextBoxes : reopened.TextBoxes;
        Assert.Equal(new[] { "First edited", "Second edited" }, boxes.Select(box => box.Paragraphs.Single().Text.TrimEnd('\n')).ToArray());
        Assert.Equal(new[] { "First description", "Second description" }, boxes.Select(box => box.Description).ToArray());
        Assert.Empty(reopened.ValidateDocument());
        string pdfText = PdfReadDocument.Open(reopened.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false })).ExtractText();
        Assert.Contains("First edited", pdfText);
        Assert.Contains("Second edited", pdfText);
    }
}
