namespace OfficeIMO.Latex.Tests;

public sealed class LatexLosslessProfileTests {
    private const string Article =
        "\\documentclass[11pt]{article}\r\n" +
        "\\usepackage{amsmath}\n" +
        "\\title{Source-Preserving Science}\r" +
        "\\author{A. Author}\r\n" +
        "% retained preamble comment\n" +
        "\\begin{document}\r\n" +
        "\\maketitle\n\n" +
        "\\section[Intro]{Introduction}\r\n" +
        "Text with \\textbf{bold} and inline math $x_1^2$.\n\n" +
        "\\begin{equation}\nE = mc^2\n\\end{equation}\r\n" +
        "\\unknowncommand{kept}\n" +
        "\\end{document}";

    [Fact]
    public void ScientificArticle_IsStructuredRecognizedAndCharacterLossless() {
        LatexParseResult result = LatexDocument.Parse(Article);
        LatexDocument document = result.Document;

        Assert.True(result.IsLossless);
        Assert.False(result.HasErrors);
        Assert.Equal(Article, document.ToLatex());
        Assert.Equal("article", document.DocumentClassName);
        Assert.True(document.IsRecognizedProfile);
        Assert.NotNull(document.Body);
        Assert.Contains(document.Environments, environment => environment.Name == "equation" && environment.IsMath);
        Assert.Equal(2, document.Math.Count);
        LatexHeading heading = Assert.Single(document.Headings);
        Assert.Equal("Introduction", heading.Title);
        Assert.Equal("Intro", heading.ShortTitle);
        Assert.Contains(document.Paragraphs, paragraph => paragraph.Content.Contains("Text with", StringComparison.Ordinal));
        LatexCommand unknown = Assert.Single(document.Commands, command => command.Name == "unknowncommand");
        Assert.False(unknown.IsProfileKnown);
        Assert.Equal("kept", unknown.GetRequiredArgument(0)!.Content);
        Assert.All(document.SyntaxTree.Root.DescendantsAndSelf(), node => AssertCoverage(node, Article));
    }

    [Theory]
    [InlineData("article", "section")]
    [InlineData("report", "chapter")]
    [InlineData("book", "chapter")]
    public void ArticleReportAndBook_ProfileShapesAreRecognized(string documentClass, string headingCommand) {
        string source = "\\documentclass{" + documentClass + "}\n\\begin{document}\n\\" + headingCommand + "{Title}\nBody.\n\\end{document}\n";

        LatexDocument document = LatexDocument.Parse(source).Document;

        Assert.True(document.IsRecognizedProfile);
        Assert.Equal("Title", Assert.Single(document.Headings).Title);
        Assert.Contains(document.Paragraphs, paragraph => paragraph.Content == "Body.");
        Assert.Equal(source, document.ToLatex());
    }

    [Fact]
    public void InlineDisplayAndCommandDelimitedMath_AreTypedAndPreserved() {
        const string source = "\\documentclass{article}\n\\begin{document}\n$a$ \\(b\\) $$c$$ \\[d\\]\n\\end{document}";

        LatexDocument document = LatexDocument.Parse(source).Document;

        Assert.Equal(new[] {
            LatexMathKind.InlineDollar,
            LatexMathKind.InlineParentheses,
            LatexMathKind.DisplayDollar,
            LatexMathKind.DisplayBrackets
        }, document.Math.Select(math => math.Kind));
        Assert.Equal(new[] { "a", "b", "c", "d" }, document.Math.Select(math => math.Content));
        Assert.Equal(source, document.ToLatex());
    }

    private static void AssertCoverage(LatexSyntaxNode node, string source) {
        Assert.Equal(node.OriginalText, node.Span.Slice(source));
        if (node.Children.Count == 0) return;
        int expected = node.Span.Start.Offset;
        foreach (LatexSyntaxNode child in node.Children) {
            Assert.Equal(expected, child.Span.Start.Offset);
            expected = child.Span.End.Offset;
        }
        Assert.Equal(node.Span.End.Offset, expected);
    }
}
