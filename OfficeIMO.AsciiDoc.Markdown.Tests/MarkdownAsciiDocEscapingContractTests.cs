namespace OfficeIMO.AsciiDoc.Markdown.Tests;

public sealed class MarkdownAsciiDocEscapingContractTests {
    [Theory]
    [InlineData("= Literal text")]
    [InlineData(".Title")]
    [InlineData("----")]
    [InlineData("include::secret[]")]
    [InlineData("NOTE: literal paragraph")]
    [InlineData(":attribute: value")]
    [InlineData("Term:: literal definition")]
    [InlineData("// literal comment")]
    [InlineData("[literal-role]")]
    [InlineData("[[literal-anchor]]")]
    public void TypedPlainParagraphsRemainLiteral(string literal) {
        MarkdownDoc source = MarkdownDoc.Create().Add(new ParagraphBlock(new InlineSequence().Text(literal)));
        MarkdownToAsciiDocResult result = source.ToAsciiDocDocumentResult();
        Assert.Single(result.Value.Blocks.OfType<AsciiDocParagraph>());
        AsciiDocToMarkdownResult reopened = result.Value.ToMarkdownDocumentResult();
        ParagraphBlock paragraph = Assert.Single(reopened.Value.Blocks.OfType<ParagraphBlock>());
        Assert.Equal(literal, string.Concat(paragraph.Inlines.Nodes.OfType<MarkdownTextRun>().Select(text => text.Text)));
        Assert.False(result.HasLoss);
        Assert.False(reopened.HasLoss);
    }

    [Theory]
    [InlineData("a**b**c")]
    [InlineData("a*b*c")]
    [InlineData("a`b`c")]
    public void IntrawordFormattingSurvivesNativeProjection(string markdown) {
        MarkdownDoc source = MarkdownReader.Parse(markdown);
        MarkdownToAsciiDocResult result = source.ToAsciiDocDocumentResult();
        Assert.Equal(source.ToMarkdown(), result.Value.ToMarkdownDocumentResult().Value.ToMarkdown());
        Assert.False(result.HasLoss);
    }

    [Fact]
    public void CodeSpanPreservesLiteralFormattingCharacters() {
        var content = new InlineSequence().Code("*literal* {attribute} \\ `code`");
        MarkdownDoc source = MarkdownDoc.Create().Add(new ParagraphBlock(content));
        MarkdownDoc reopened = source.ToAsciiDocDocumentResult().Value.ToMarkdownDocumentResult().Value;
        CodeSpanInline code = Assert.Single(Assert.Single(reopened.Blocks.OfType<ParagraphBlock>()).Inlines.Nodes.OfType<CodeSpanInline>());
        Assert.Equal("*literal* {attribute} \\ `code`", code.Text);
    }
}
