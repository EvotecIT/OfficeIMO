using System.Collections.Generic;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.MarkdownRenderer;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_Semantic_Fenced_Block_Ast_Parity_Tests {
    [Fact]
    public void Semantic_Fenced_Block_Ast_Parity_Holds_For_Rendered_Chart_Block() {
        const string payload = "{\r\n  \"type\": \"bar\",\r\n  \"data\": { \"labels\": [\"A\"], \"datasets\": [{ \"label\": \"Count\", \"data\": [1] }] }\r\n}";
        string markdown = "```ix-chart\n" + payload.Replace("\r\n", "\n") + "\n```";
        string html = MarkdownVisualContract.BuildElementHtml(
            "canvas",
            "omd-visual omd-chart",
            MarkdownSemanticKinds.Chart,
            "ix-chart",
            MarkdownVisualContract.CreatePayload(payload));

        AssertSemanticBlockParity(
            MarkdownReader.Parse(markdown, CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart)),
            html.LoadFromHtml());
    }

    [Fact]
    public void Semantic_Fenced_Block_Ast_Parity_Holds_For_Rendered_Mermaid_Block() {
        const string payload = "flowchart LR\r\nA-->B\r\nB-->C";
        string markdown = "```mermaid\n" + payload.Replace("\r\n", "\n") + "\n```";
        string html = "<pre class=\"mermaid\">" + System.Net.WebUtility.HtmlEncode(payload) + "</pre>";

        AssertSemanticBlockParity(
            MarkdownReader.Parse(markdown, CreateSemanticOptions("mermaid", MarkdownSemanticKinds.Mermaid)),
            html.LoadFromHtml());
    }

    [Fact]
    public void Semantic_Fenced_Block_Ast_Parity_Holds_For_Rendered_Math_Block() {
        const string payload = "x^2 + 1\r\ny = 2";
        string markdown = "```math\n" + payload.Replace("\r\n", "\n") + "\n```";
        string html = "<div class=\"omd-math\">$$\r\n" + System.Net.WebUtility.HtmlEncode(payload) + "\r\n$$</div>";

        AssertSemanticBlockParity(
            MarkdownReader.Parse(markdown, CreateSemanticOptions("math", MarkdownSemanticKinds.Math)),
            html.LoadFromHtml());
    }

    private static void AssertSemanticBlockParity(MarkdownDoc markdownDocument, MarkdownDoc htmlDocument) {
        var markdownBlock = Assert.IsType<SemanticFencedBlock>(Assert.Single(markdownDocument.Blocks));
        var htmlBlock = Assert.IsType<SemanticFencedBlock>(Assert.Single(htmlDocument.Blocks));

        Assert.Equal(markdownBlock.SemanticKind, htmlBlock.SemanticKind);
        Assert.Equal(markdownBlock.Language, htmlBlock.Language);
        Assert.Equal(markdownBlock.Content, htmlBlock.Content);
        Assert.Equal(markdownBlock.Caption, htmlBlock.Caption);
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
