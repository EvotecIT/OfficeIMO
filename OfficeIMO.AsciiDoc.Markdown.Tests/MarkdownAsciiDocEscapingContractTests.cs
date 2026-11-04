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
        MarkdownDocument source = MarkdownDocument.Create();
        source.Add(new ParagraphBlock(new InlineSequence().Text(literal)));
        var result = source.ToAsciiDoc(null, out OfficeIMO.OfficeConversionReport conversionReport1);
        var resultReport = (MarkdownToAsciiDocReport)conversionReport1;
        Assert.Single(result.Blocks.OfType<AsciiDocParagraph>());
        var reopened = result.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport2);
        var reopenedReport = (AsciiDocToMarkdownReport)conversionReport2;
        ParagraphBlock paragraph = Assert.Single(reopened.Blocks.OfType<ParagraphBlock>());
        Assert.Equal(literal, string.Concat(paragraph.Inlines.Nodes.OfType<MarkdownTextRun>().Select(text => text.Text)));
        Assert.False(resultReport.HasLoss);
        Assert.False(reopenedReport.HasLoss);
    }

    [Theory]
    [InlineData("a**b**c")]
    [InlineData("a*b*c")]
    [InlineData("a`b`c")]
    public void IntrawordFormattingSurvivesNativeProjection(string markdown) {
        MarkdownDocument source = MarkdownDocument.Parse(markdown);
        var result = source.ToAsciiDoc(null, out OfficeIMO.OfficeConversionReport conversionReport3);
        var resultReport = (MarkdownToAsciiDocReport)conversionReport3;
        Assert.Equal(source.ToMarkdown(), result.ToMarkdown(null, out OfficeIMO.OfficeConversionReport conversionReport4).ToMarkdown());
        Assert.False(resultReport.HasLoss);
    }

    [Fact]
    public void CodeSpanPreservesLiteralFormattingCharacters() {
        var content = new InlineSequence().Code("*literal* {attribute} \\ `code`");
        MarkdownDocument source = MarkdownDocument.Create();
        source.Add(new ParagraphBlock(content));
        MarkdownDocument reopened = source.ToAsciiDoc().ToMarkdown();
        CodeSpanInline code = Assert.Single(Assert.Single(reopened.Blocks.OfType<ParagraphBlock>()).Inlines.Nodes.OfType<CodeSpanInline>());
        Assert.Equal("*literal* {attribute} \\ `code`", code.Text);
    }
}
