using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class WordPageBreakConversionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WordPageBreaksSurviveTheOdtRoundTrip(bool insideParagraph) {
        using var source = WordDocument.Create();
        var first = source.AddParagraph("Before the break");
        if (insideParagraph) first.AddBreak(WordBreakType.Page).AddText("After the break");
        else { source.AddPageBreak(); source.AddParagraph("After the break"); }
        var odt = source.ToOpenDocumentResult().Value;
        using var converted = odt.ToWordDocumentResult().Value;
        Assert.Equal(2, converted.GetEstimatedPageCount());
        Assert.Contains(converted.Paragraphs, paragraph => paragraph.Text.Contains("Before the break"));
        Assert.Contains(converted.Paragraphs, paragraph => paragraph.Text.Contains("After the break"));
    }

    [Fact]
    public void ConsecutivePageBreaksPreserveTheBlankPageAndRunFormatting() {
        using var source = WordDocument.Create();
        var paragraph = source.AddParagraph("Before");
        paragraph.Bold = true;
        paragraph.AddBreak(WordBreakType.Page).AddBreak(WordBreakType.Page).AddText("After").Bold = true;
        using var converted = source.ToOpenDocumentResult().Value.ToWordDocumentResult().Value;
        Assert.Equal(3, converted.GetEstimatedPageCount());
        Assert.Contains(converted.Paragraphs, item => item.Text == "Before" && item.Bold == true);
        Assert.Contains(converted.Paragraphs, item => item.Text == "After" && item.Bold == true);
    }

    [Fact]
    public void LiteralLineSeparatorTextDoesNotCreateAPageBreak() {
        using var source = WordDocument.Create();
        source.AddParagraph("Before\u2028After");
        using var converted = source.ToOpenDocumentResult().Value.ToWordDocumentResult().Value;
        Assert.Equal(1, converted.GetEstimatedPageCount());
        Assert.Contains(converted.Paragraphs, item => item.Text.Contains("Before") && item.Text.Contains("After"));
    }
}
