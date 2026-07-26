using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlWordGapClosure {
    [Theory]
    [InlineData("""<html style="font-size:32px"><body><div style="margin-inline-start:1rem">Text</div></body></html>""")]
    [InlineData("""<style>html { font-size:32px; }</style><div style="margin-inline-start:1rem">Text</div>""")]
    public void HtmlToWord_RemLogicalSpacingUsesTheComputedRootFontSize(string html) {
        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Text");

        Assert.Equal(480, paragraph.IndentationBefore);
    }

    [Fact]
    public void HtmlToWord_EmLogicalSpacingUsesTheComputedElementFontSize() {
        const string html = """
            <div style="font-size:20px;margin-inline-start:1em">Text</div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Text");

        Assert.Equal(300, paragraph.IndentationBefore);
    }

    [Fact]
    public void HtmlToWord_InheritedEmLogicalSpacingUsesTheParentsComputedFontSizeOnce() {
        const string html = """
            <style>.parent { font-size:2em; }</style>
            <div class="parent">
              <div style="margin-inline-start:1em">Text</div>
            </div>
            """;

        using WordDocument document = HtmlConversionDocument.Parse(html).ToWordDocument();
        WordParagraph paragraph = Assert.Single(
            document.Paragraphs,
            candidate => candidate.Text == "Text");

        Assert.Equal(480, paragraph.IndentationBefore);
    }
}
