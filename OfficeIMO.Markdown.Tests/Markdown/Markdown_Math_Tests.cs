using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Math_Tests {
    private const string Source = "Energy $E = mc^2$ and $$\\sum_i x_i$$ cost $5 and $10.\n\n$$\n\\frac{a}{b}\n$$\n\n```math\n\\sqrt{x}\n```";

    [Fact]
    public void Math_Delimiters_Stay_Text_Unless_The_Profile_Is_Enabled() {
        var document = MarkdownDocument.Parse(Source);

        Assert.DoesNotContain(document.Blocks, block => block is MathBlock || block is SemanticFencedBlock);
        Assert.DoesNotContain(((ParagraphBlock)document.Blocks[0]).Inlines.Nodes, inline => inline is MathInline);
    }

    [Fact]
    public void Math_Profile_Produces_Math_Nodes_With_Source_Spans() {
        var options = new MarkdownLoadOptions { Math = true };
        var result = MarkdownReader.ParseWithSyntaxTree(Source, options);
        MarkdownInvariantAssert.SyntaxTreeIsWellFormed(result.FinalSyntaxTree);
        var document = result.Document;

        var paragraph = Assert.IsType<ParagraphBlock>(document.Blocks[0]);
        var math = paragraph.Inlines.Nodes.OfType<MathInline>().ToArray();
        Assert.Equal(2, math.Length);
        Assert.Equal("E = mc^2", math[0].Source);
        Assert.False(math[0].IsDisplay);
        Assert.Equal(new MarkdownSourceSpan(1, 9, 1, 16), StripOffsets(math[0].ContentSourceSpan));
        Assert.Equal("\\sum_i x_i", math[1].Source);
        Assert.True(math[1].IsDisplay);
        Assert.Contains("cost $5 and $10.", paragraph.Inlines.RenderMarkdown());

        var display = Assert.IsType<MathBlock>(document.Blocks[1]);
        Assert.Equal("\\frac{a}{b}", display.Content);
        Assert.Equal(new MarkdownSourceSpan(4, 1, 4, 11), StripOffsets(display.ContentSourceSpan));
        Assert.Equal(3, display.OpeningDelimiterSourceSpan!.Value.StartLine);
        Assert.Equal(5, display.ClosingDelimiterSourceSpan!.Value.StartLine);

        var fence = Assert.IsType<SemanticFencedBlock>(document.Blocks[2]);
        Assert.Equal("math", fence.SemanticKind);
        Assert.Equal("\\sqrt{x}", fence.Content);

        Assert.Contains(result.FinalSyntaxTree.Descendants(), node => node.Kind == MarkdownSyntaxKind.InlineMath);
        Assert.Contains(result.FinalSyntaxTree.Descendants(), node => node.Kind == MarkdownSyntaxKind.MathBlock);
    }

    [Fact]
    public void Math_Profile_Writes_Markdown_That_Reparses_To_The_Same_Math() {
        var options = new MarkdownLoadOptions { Math = true };
        var document = MarkdownDocument.Parse(Source, options);

        string written = document.ToMarkdown();
        var reparsed = MarkdownDocument.Parse(written, new MarkdownLoadOptions { Math = true });

        Assert.Equal(Describe(document), Describe(reparsed));
        Assert.Contains("$$\n\\frac{a}{b}\n$$", written.Replace("\r\n", "\n"));
    }

    [Theory]
    [InlineData("Pay $5 and $10 today.")]
    [InlineData("A lone $ sign.")]
    [InlineData("Escaped \\$x\\$ stays text.")]
    [InlineData("Code `$x$` stays code.")]
    public void Math_Profile_Leaves_Non_Math_Dollars_Alone(string markdown) {
        var document = MarkdownDocument.Parse(markdown, new MarkdownLoadOptions { Math = true });

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        Assert.DoesNotContain(paragraph.Inlines.Nodes, inline => inline is MathInline);
    }

    private static string Describe(MarkdownDocument document) =>
        string.Join("|", document.DescendantsAndSelf().SelectMany(block => block switch {
            ParagraphBlock paragraph => paragraph.Inlines.Nodes.OfType<MathInline>().Select(inline => "inline:" + inline.IsDisplay + ":" + inline.Source),
            MathBlock math => new[] { "block:" + math.Content },
            SemanticFencedBlock fence => new[] { "fence:" + fence.SemanticKind + ":" + fence.Content },
            _ => Array.Empty<string>()
        }));

    private static MarkdownSourceSpan? StripOffsets(MarkdownSourceSpan? span) =>
        span.HasValue
            ? new MarkdownSourceSpan(span.Value.StartLine, span.Value.StartColumn ?? 0, span.Value.EndLine, span.Value.EndColumn ?? 0)
            : null;
}
