using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("Consolas", null, true)]
    [InlineData("Courier New", null, true)]
    [InlineData("Calibri", null, false)]
    [InlineData("Calibri", "HTMLCode", true)]
    public void GetFormattedRuns_DetectsCodeByFamilyOrStyleOnAnyMachine(string family, string? characterStyle, bool expected) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Body ");
        WordParagraph run = paragraph.AddText("value");
        run.SetFontFamily(family);
        if (characterStyle != null) run.CharacterStyleId = characterStyle;

        WordFormattedRun formatted = Assert.Single(paragraph.GetFormattedRuns(), item => item.Text == "value");
        Assert.Equal(expected, formatted.Code);
    }

    [Fact]
    public void GetFormattedRuns_DocumentBodyFontIsNotCode() {
        using WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = "Consolas";
        WordParagraph paragraph = document.AddParagraph("Body ");
        paragraph.AddText("value").SetFontFamily("Consolas");
        Assert.DoesNotContain(paragraph.GetFormattedRuns(), item => item.Code);
    }

    [Fact]
    public void MarkdownToWord_InlineCodeIsSetSmallerThanBodyText() {
        using WordDocument document = OfficeIMO.Markdown.MarkdownDocument.Parse("Run `dotnet test` now.").ToWord();
        WordParagraph code = document.Paragraphs.Single(paragraph => paragraph.Text == "dotnet test");
        int body = document.Settings.FontSize ?? 11;
        Assert.NotNull(code.FontSize);
        Assert.True(code.FontSize < body, $"Inline code is {code.FontSize} pt on {body} pt body text.");
    }
}
