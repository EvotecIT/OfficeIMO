using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Fonts the converters write into documents are the authored names, never a substitute for a font the generating
/// machine lacks, so a document comes out the same on every platform.
/// </summary>
public sealed class WordDocumentFontTests {
    // No machine has this face, so a lookup of installed fonts would replace it everywhere.
    private const string AbsentFont = "OfficeIMO Absent Test Face";

    [Theory]
    [InlineData("serif", "Times New Roman")]
    [InlineData("SANS-SERIF", "Calibri")]
    [InlineData(" monospace ", "Consolas")]
    [InlineData("cursive", "Comic Sans MS")]
    [InlineData("fantasy", "Impact")]
    [InlineData(" " + AbsentFont + " ", AbsentFont)]
    public void DocumentFontsMapGenericFamiliesAndKeepAuthoredNames(string authored, string expected) {
        Assert.Equal(expected, WordFontResolver.ToDocumentFont(authored));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void AnEmptyDocumentFontIsNull(string? authored) {
        Assert.Null(WordFontResolver.ToDocumentFont(authored));
    }

    [Fact]
    public void MarkdownToWordWritesTheAuthoredFontFamily() {
        using MarkdownDocument source = MarkdownDocument.Parse("Body text.");
        using WordDocument document = source.ToWord(new MarkdownWordExportOptions { FontFamily = AbsentFont });

        Assert.Equal(AbsentFont, RunWith(document, "Body").FontFamily);
    }

    [Fact]
    public void HtmlToWordWritesTheFirstAuthoredFamilyAndMapsGenericOnes() {
        string html = $"<p style=\"font-family:'{AbsentFont}', serif\">Authored</p><p style=\"font-family:serif\">Generic</p>";
        using WordDocument document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWord();

        Assert.Equal(AbsentFont, RunWith(document, "Authored").FontFamily);
        Assert.Equal("Times New Roman", RunWith(document, "Generic").FontFamily);
    }

    [Fact]
    public void WordToMarkdownRecognizesTheConfiguredCodeFontAsAuthored() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Run ");
        paragraph.AddText("dotnet test").SetFontFamily(AbsentFont);

        string markdown = document.ToMarkdown(new WordMarkdownExportOptions { FontFamily = AbsentFont }).ToMarkdown();

        Assert.Contains("`dotnet test`", markdown);
    }

    private static WordParagraph RunWith(WordDocument document, string text) =>
        document.Paragraphs.SelectMany(paragraph => paragraph.GetRuns()).First(run => run.Text.Contains(text));
}
