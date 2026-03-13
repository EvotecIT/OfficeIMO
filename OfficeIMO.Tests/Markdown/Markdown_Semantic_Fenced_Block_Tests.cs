using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Semantic_Fenced_Block_Tests {
    [Fact]
    public void Parse_Uses_FencedBlockExtension_To_Create_SemanticFencedBlock() {
        var options = CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart);
        var markdown = """
```ix-chart
{"type":"bar"}
```
_Chart caption_
""";

        var doc = MarkdownReader.Parse(markdown, options);

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(doc.Blocks));
        Assert.Equal(MarkdownSemanticKinds.Chart, block.SemanticKind);
        Assert.Equal("ix-chart", block.Language);
        Assert.Equal("{\"type\":\"bar\"}", block.Content);
        Assert.Equal("Chart caption", block.Caption);
    }

    [Fact]
    public void Parse_Uses_FencedBlockExtension_For_Nested_Quoted_Fences() {
        var options = CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart);
        var markdown = """
> ```ix-chart
> {"type":"bar"}
> ```
""";

        var doc = MarkdownReader.Parse(markdown, options);

        var quote = Assert.IsType<QuoteBlock>(Assert.Single(doc.Blocks));
        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(quote.ChildBlocks));
        Assert.Equal(MarkdownSemanticKinds.Chart, block.SemanticKind);
        Assert.Equal("ix-chart", block.Language);
    }

    [Fact]
    public void ParseWithSyntaxTree_Captures_SemanticFencedBlock_Structure() {
        var options = CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart);
        var markdown = """
```ix-chart
{"type":"bar"}
```
""";

        var result = MarkdownReader.ParseWithSyntaxTree(markdown, options);

        var block = Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(MarkdownSyntaxKind.SemanticFencedBlock, block.Kind);
        Assert.Equal(3, block.Children.Count);
        Assert.Equal(MarkdownSyntaxKind.FenceSemanticKind, block.Children[0].Kind);
        Assert.Equal(MarkdownSemanticKinds.Chart, block.Children[0].Literal);
        Assert.Equal(MarkdownSyntaxKind.CodeFenceInfo, block.Children[1].Kind);
        Assert.Equal("ix-chart", block.Children[1].Literal);
        Assert.Equal(MarkdownSyntaxKind.CodeContent, block.Children[2].Kind);
        Assert.Equal("{\"type\":\"bar\"}", block.Children[2].Literal);
    }

    [Fact]
    public void SemanticFencedBlock_RenderHtml_Uses_SemanticRenderer_And_RoundTrips_Markdown() {
        var block = new SemanticFencedBlock(MarkdownSemanticKinds.Chart, "ix-chart", "{\"type\":\"bar\"}", "Chart caption");
        var doc = MarkdownDoc.Create().Add(block);
        var html = doc.ToHtmlFragment(new HtmlOptions {
            Kind = HtmlKind.Fragment,
            SemanticFencedBlockHtmlRenderer = static (semanticBlock, _) =>
                $"<div class=\"semantic-block\" data-kind=\"{semanticBlock.SemanticKind}\" data-language=\"{semanticBlock.Language}\"></div>"
        });

        Assert.Contains("class=\"semantic-block\"", html, StringComparison.Ordinal);
        Assert.Contains($"data-kind=\"{MarkdownSemanticKinds.Chart}\"", html, StringComparison.Ordinal);
        Assert.Contains("data-language=\"ix-chart\"", html, StringComparison.Ordinal);
        var markdown = ((IMarkdownBlock)block).RenderMarkdown().Replace("\r\n", "\n");
        Assert.Equal("```ix-chart\n{\"type\":\"bar\"}\n```\n_Chart caption_", markdown);
    }

    [Fact]
    public void SemanticFencedBlock_RenderHtml_Falls_Back_To_CodeBlockHtmlRenderer_When_Needed() {
        var block = new SemanticFencedBlock("note", "ix-note", "hello");
        var doc = MarkdownDoc.Create().Add(block);
        var html = doc.ToHtmlFragment(new HtmlOptions {
            Kind = HtmlKind.Fragment,
            CodeBlockHtmlRenderer = static (codeBlock, _) =>
                $"<aside class=\"code-fallback\" data-language=\"{codeBlock.Language}\">{System.Net.WebUtility.HtmlEncode(codeBlock.Content)}</aside>"
        });

        Assert.Contains("class=\"code-fallback\"", html, StringComparison.Ordinal);
        Assert.Contains("data-language=\"ix-note\"", html, StringComparison.Ordinal);
        Assert.Contains(">hello<", html, StringComparison.Ordinal);
    }

    private static MarkdownReaderOptions CreateSemanticOptions(string language, string semanticKind) {
        var options = new MarkdownReaderOptions();
        options.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
            "Semantic AST",
            new[] { language },
            context => new SemanticFencedBlock(semanticKind, context.Language, context.Content, context.Caption)));
        return options;
    }
}
