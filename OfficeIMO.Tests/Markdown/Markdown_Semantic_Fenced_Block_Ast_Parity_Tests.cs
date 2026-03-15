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

    [Fact]
    public void Parse_Block_Fragment_Preserves_Semantic_Fenced_Block_Extensions() {
        const string payload = "{\n  \"type\": \"bar\"\n}";
        string markdown = "```ix-chart\n" + payload + "\n```";

        var blocks = MarkdownReader.ParseBlockFragment(markdown, CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart));

        var block = Assert.IsType<SemanticFencedBlock>(Assert.Single(blocks));
        Assert.Equal("ix-chart", block.Language);
        Assert.Equal(payload, block.Content);
    }

    [Fact]
    public void Registered_Fenced_Block_Transform_Rewrites_Nested_List_CodeBlocks() {
        const string payload = "{\"type\":\"bar\"}";
        var options = CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart);
        var transform = new MarkdownRegisteredFencedBlockTransform(options.FencedBlockExtensions);
        var item = ListItem.Text("Intro");
        item.Children.Add(new CodeBlock("ix-chart", payload));
        var document = MarkdownDoc.Create().Add(new UnorderedListBlock {
            Items = { item }
        });

        transform.Transform(
            document,
            new MarkdownDocumentTransformContext(MarkdownDocumentTransformSource.MarkdownReader, options));

        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        var semantic = Assert.IsType<SemanticFencedBlock>(Assert.Single(list.Items[0].Children));
        Assert.Equal(payload, semantic.Content);
    }

    [Fact]
    public void Semantic_Fenced_Block_Ast_Parity_Holds_Inside_List_Items() {
        const string payload = "{\r\n  \"type\": \"bar\",\r\n  \"data\": { \"labels\": [\"A\"], \"datasets\": [{ \"label\": \"Count\", \"data\": [1] }] }\r\n}";
        string nestedPayload = payload.Replace("\r\n", "\n").Replace("\n", "\n  ");
        string markdown = "- Intro\n\n  ```ix-chart\n"
            + "  "
            + nestedPayload
            + "\n  ```\n\n  Tail\n";
        string html = "<ul><li><p>Intro</p>"
            + MarkdownVisualContract.BuildElementHtml(
                "canvas",
                "omd-visual omd-chart",
                MarkdownSemanticKinds.Chart,
                "ix-chart",
                MarkdownVisualContract.CreatePayload(payload))
            + "<p>Tail</p></li></ul>";

        var markdownDocument = MarkdownReader.Parse(markdown, CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart));
        var htmlDocument = html.LoadFromHtml();

        var markdownList = Assert.IsType<UnorderedListBlock>(Assert.Single(markdownDocument.Blocks));
        var htmlList = Assert.IsType<UnorderedListBlock>(Assert.Single(htmlDocument.Blocks));

        var markdownItem = Assert.Single(markdownList.Items);
        var htmlItem = Assert.Single(htmlList.Items);

        Assert.Equal(markdownItem.Content.RenderMarkdown(), htmlItem.Content.RenderMarkdown());
        Assert.Equal(DescribeBlocks(markdownItem.Children), DescribeBlocks(htmlItem.Children));
    }

    [Fact]
    public void Semantic_Fenced_Block_Ast_Parity_Holds_Inside_Details_Blocks() {
        const string payload = "{\r\n  \"type\": \"bar\",\r\n  \"data\": { \"labels\": [\"A\"], \"datasets\": [{ \"label\": \"Count\", \"data\": [1] }] }\r\n}";
        string markdown = "<details open>\n<summary>More</summary>\n\n```ix-chart\n"
            + payload.Replace("\r\n", "\n")
            + "\n```\n\nTail\n</details>\n";
        string html = "<details open><summary>More</summary>"
            + MarkdownVisualContract.BuildElementHtml(
                "canvas",
                "omd-visual omd-chart",
                MarkdownSemanticKinds.Chart,
                "ix-chart",
                MarkdownVisualContract.CreatePayload(payload))
            + "<p>Tail</p></details>";

        var markdownDocument = MarkdownReader.Parse(markdown, CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart));
        var htmlDocument = html.LoadFromHtml();

        var markdownDetails = Assert.IsType<DetailsBlock>(Assert.Single(markdownDocument.Blocks));
        var htmlDetails = Assert.IsType<DetailsBlock>(Assert.Single(htmlDocument.Blocks));

        Assert.Equal(markdownDetails.Summary!.Inlines.RenderMarkdown(), htmlDetails.Summary!.Inlines.RenderMarkdown());
        Assert.Equal(DescribeBlocks(markdownDetails.ChildBlocks), DescribeBlocks(htmlDetails.ChildBlocks));
    }

    [Fact]
    public void Semantic_Fenced_Block_Ast_Parity_Holds_Inside_Table_Cells() {
        const string payload = "{\r\n  \"type\": \"bar\",\r\n  \"data\": { \"labels\": [\"A\"], \"datasets\": [{ \"label\": \"Count\", \"data\": [1] }] }\r\n}";
        string cellPayload = payload.Replace("\r\n", "<br>").Replace('\r', '\n').Replace("\n", "<br>");
        string markdown = "| Notes |\n| --- |\n| Intro<br><br>```ix-chart<br>"
            + cellPayload
            + "<br>```<br><br>Tail |\n";
        string html = "<table><tr><th>Notes</th></tr><tr><td><p>Intro</p>"
            + MarkdownVisualContract.BuildElementHtml(
                "canvas",
                "omd-visual omd-chart",
                MarkdownSemanticKinds.Chart,
                "ix-chart",
                MarkdownVisualContract.CreatePayload(payload))
            + "<p>Tail</p></td></tr></table>";

        var markdownDocument = MarkdownReader.Parse(markdown, CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart));
        var htmlDocument = html.LoadFromHtml();

        var markdownTable = Assert.IsType<TableBlock>(Assert.Single(markdownDocument.Blocks));
        var htmlTable = Assert.IsType<TableBlock>(Assert.Single(htmlDocument.Blocks));

        Assert.Equal(
            DescribeBlocks(markdownTable.RowCells[0][0].Blocks),
            DescribeBlocks(htmlTable.RowCells[0][0].Blocks));
    }

    private static void AssertSemanticBlockParity(MarkdownDoc markdownDocument, MarkdownDoc htmlDocument) {
        var markdownBlock = Assert.IsType<SemanticFencedBlock>(Assert.Single(markdownDocument.Blocks));
        var htmlBlock = Assert.IsType<SemanticFencedBlock>(Assert.Single(htmlDocument.Blocks));

        Assert.Equal(markdownBlock.SemanticKind, htmlBlock.SemanticKind);
        Assert.Equal(markdownBlock.Language, htmlBlock.Language);
        Assert.Equal(markdownBlock.Content, htmlBlock.Content);
        Assert.Equal(markdownBlock.Caption, htmlBlock.Caption);
    }

    private static string DescribeBlocks(IReadOnlyList<IMarkdownBlock> blocks) {
        var parts = new List<string>(blocks.Count);
        for (int i = 0; i < blocks.Count; i++) {
            parts.Add(DescribeBlock(blocks[i]));
        }

        return string.Join(" | ", parts);
    }

    private static string DescribeBlock(IMarkdownBlock block) {
        return block switch {
            ParagraphBlock paragraph => "Paragraph:" + paragraph.Inlines.RenderMarkdown(),
            SemanticFencedBlock semantic => "Semantic:" + semantic.SemanticKind + ":" + semantic.Language + ":" + semantic.Content,
            _ => block.GetType().Name + ":" + block.RenderMarkdown()
        };
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
