using System.Collections.Generic;
using System.Linq;
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
    public void Semantic_Fenced_Block_Ast_Parity_Holds_For_Rendered_Chart_Block_With_Fence_Metadata() {
        const string payload = "{\r\n  \"type\": \"bar\",\r\n  \"data\": { \"labels\": [\"A\"], \"datasets\": [{ \"label\": \"Count\", \"data\": [1] }] }\r\n}";
        string markdown = "```ix-chart #quarterly-summary .wide title=\"Quarterly Overview\" pinned\n" + payload.Replace("\r\n", "\n") + "\n```";
        string html = MarkdownVisualContract.BuildElementHtml(
            "canvas",
            "omd-visual omd-chart",
            MarkdownSemanticKinds.Chart,
            "ix-chart",
            MarkdownVisualContract.CreatePayload(payload),
            MarkdownCodeFenceInfo.Parse("ix-chart #quarterly-summary .wide title=\"Quarterly Overview\" pinned"));

        AssertSemanticBlockParity(
            MarkdownReader.Parse(markdown, CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart)),
            html.LoadFromHtml());
    }

    [Fact]
    public void Semantic_Fenced_Block_Ast_Parity_Holds_For_Rendered_Chart_Block_With_Opaque_Fence_Metadata_Tail() {
        const string payload = "{\r\n  \"type\": \"bar\",\r\n  \"data\": { \"labels\": [\"A\"], \"datasets\": [{ \"label\": \"Count\", \"data\": [1] }] }\r\n}";
        string markdown = "```ix-chart {#quarterly-summary .wide title=\"Quarterly Overview\"\n" + payload.Replace("\r\n", "\n") + "\n```";
        string html = MarkdownVisualContract.BuildElementHtml(
            "canvas",
            "omd-visual omd-chart",
            MarkdownSemanticKinds.Chart,
            "ix-chart",
            MarkdownVisualContract.CreatePayload(payload),
            MarkdownCodeFenceInfo.Parse("ix-chart {#quarterly-summary .wide title=\"Quarterly Overview\""));

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
    public void Registered_Fenced_Block_Transform_Skips_Rewrite_When_No_Matching_CodeBlock_Exists() {
        var options = CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart);
        var transform = new MarkdownRegisteredFencedBlockTransform(options.FencedBlockExtensions);
        var lead = new InlineSequence().Text("Intro");
        var additional = new InlineSequence().Text("Tail");
        var item = new ListItem(lead);
        item.AdditionalParagraphs.Add(additional);
        var document = MarkdownDoc.Create().Add(new UnorderedListBlock {
            Items = { item }
        });

        transform.Transform(
            document,
            new MarkdownDocumentTransformContext(MarkdownDocumentTransformSource.MarkdownReader, options));

        var list = Assert.IsType<UnorderedListBlock>(Assert.Single(document.Blocks));
        Assert.Same(lead, list.Items[0].Content);
        Assert.Same(additional, list.Items[0].AdditionalParagraphs[0]);
        Assert.Equal(new[] { "Intro", "Tail" }, list.Items[0].BlockChildren.Select(block => block.RenderMarkdown()).ToArray());
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
        Assert.Equal(MarkdownAstParityFormatter.DescribeBlocks(markdownItem.Children), MarkdownAstParityFormatter.DescribeBlocks(htmlItem.Children));
    }

    [Fact]
    public void Semantic_Fenced_Block_Ast_Parity_Holds_Inside_List_Items_With_Details() {
        const string payload = "{\r\n  \"type\": \"bar\",\r\n  \"data\": { \"labels\": [\"A\"], \"datasets\": [{ \"label\": \"Count\", \"data\": [1] }] }\r\n}";
        string nestedPayload = payload.Replace("\r\n", "\n").Replace("\n", "\n  ");
        string markdown = "- Intro\n\n  <details open>\n  <summary>More</summary>\n\n  Hidden\n\n  ```ix-chart\n"
            + "  "
            + nestedPayload
            + "\n  ```\n\n  </details>\n\n  Tail\n";
        string html = "<ul><li><p>Intro</p><details open><summary>More</summary><p>Hidden</p>"
            + MarkdownVisualContract.BuildElementHtml(
                "canvas",
                "omd-visual omd-chart",
                MarkdownSemanticKinds.Chart,
                "ix-chart",
                MarkdownVisualContract.CreatePayload(payload))
            + "</details><p>Tail</p></li></ul>";

        var markdownDocument = MarkdownReader.Parse(markdown, CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart));
        var htmlDocument = html.LoadFromHtml();

        var markdownList = Assert.IsType<UnorderedListBlock>(Assert.Single(markdownDocument.Blocks));
        var htmlList = Assert.IsType<UnorderedListBlock>(Assert.Single(htmlDocument.Blocks));

        var markdownItem = Assert.Single(markdownList.Items);
        var htmlItem = Assert.Single(htmlList.Items);

        Assert.Equal(markdownItem.Content.RenderMarkdown(), htmlItem.Content.RenderMarkdown());
        Assert.Equal(MarkdownAstParityFormatter.DescribeBlocks(markdownItem.Children), MarkdownAstParityFormatter.DescribeBlocks(htmlItem.Children));
    }

    [Fact]
    public void Semantic_Fenced_Block_Ast_Parity_Holds_Inside_List_Items_With_Details_And_Callout() {
        const string payload = "{\r\n  \"type\": \"bar\",\r\n  \"data\": { \"labels\": [\"A\"], \"datasets\": [{ \"label\": \"Count\", \"data\": [1] }] }\r\n}";
        string nestedPayload = payload.Replace("\r\n", "\n").Replace("\n", "\n  > ");
        string markdown = "- Intro\n\n  <details open>\n  <summary>More</summary>\n\n  > [!NOTE] Watch\n  > Body\n  >\n  > ```ix-chart\n  > "
            + nestedPayload
            + "\n  > ```\n\n  </details>\n";
        string html = "<ul><li><p>Intro</p><details open><summary>More</summary><blockquote class=\"callout note\"><p><strong>Watch</strong></p><p>Body</p>"
            + MarkdownVisualContract.BuildElementHtml(
                "canvas",
                "omd-visual omd-chart",
                MarkdownSemanticKinds.Chart,
                "ix-chart",
                MarkdownVisualContract.CreatePayload(payload))
            + "</blockquote></details></li></ul>";

        var markdownDocument = MarkdownReader.Parse(markdown, CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart));
        var htmlDocument = html.LoadFromHtml();

        var markdownList = Assert.IsType<UnorderedListBlock>(Assert.Single(markdownDocument.Blocks));
        var htmlList = Assert.IsType<UnorderedListBlock>(Assert.Single(htmlDocument.Blocks));

        var markdownItem = Assert.Single(markdownList.Items);
        var htmlItem = Assert.Single(htmlList.Items);

        Assert.Equal(markdownItem.Content.RenderMarkdown(), htmlItem.Content.RenderMarkdown());
        Assert.Equal(MarkdownAstParityFormatter.DescribeBlocks(markdownItem.Children), MarkdownAstParityFormatter.DescribeBlocks(htmlItem.Children));
    }

    [Fact]
    public void Semantic_Fenced_Block_Ast_Parity_Holds_Inside_Ordered_List_Items() {
        const string payload = "{\r\n  \"type\": \"bar\",\r\n  \"data\": { \"labels\": [\"A\"], \"datasets\": [{ \"label\": \"Count\", \"data\": [1] }] }\r\n}";
        string nestedPayload = payload.Replace("\r\n", "\n").Replace("\n", "\n   ");
        string markdown = "1. Intro\n\n   ```ix-chart\n"
            + "   "
            + nestedPayload
            + "\n   ```\n\n   Tail\n";
        string html = "<ol><li><p>Intro</p>"
            + MarkdownVisualContract.BuildElementHtml(
                "canvas",
                "omd-visual omd-chart",
                MarkdownSemanticKinds.Chart,
                "ix-chart",
                MarkdownVisualContract.CreatePayload(payload))
            + "<p>Tail</p></li></ol>";

        var markdownDocument = MarkdownReader.Parse(markdown, CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart));
        var htmlDocument = html.LoadFromHtml();

        var markdownList = Assert.IsType<OrderedListBlock>(Assert.Single(markdownDocument.Blocks));
        var htmlList = Assert.IsType<OrderedListBlock>(Assert.Single(htmlDocument.Blocks));

        var markdownItem = Assert.Single(markdownList.Items);
        var htmlItem = Assert.Single(htmlList.Items);

        Assert.Equal(markdownItem.Content.RenderMarkdown(), htmlItem.Content.RenderMarkdown());
        Assert.Equal(MarkdownAstParityFormatter.DescribeBlocks(markdownItem.Children), MarkdownAstParityFormatter.DescribeBlocks(htmlItem.Children));
    }

    [Fact]
    public void Semantic_Fenced_Block_Ast_Parity_Holds_Inside_Ordered_Task_List_Items_With_Details_And_Callout() {
        const string payload = "{\r\n  \"type\": \"bar\",\r\n  \"data\": { \"labels\": [\"A\"], \"datasets\": [{ \"label\": \"Count\", \"data\": [1] }] }\r\n}";
        string nestedPayload = payload.Replace("\r\n", "\n").Replace("\n", "\n   > ");
        string markdown = "1. [x] Check\n\n   <details open>\n   <summary>More</summary>\n\n   > [!NOTE] Watch\n   > Body\n   >\n   > ```ix-chart\n   > "
            + nestedPayload
            + "\n   > ```\n\n   </details>\n";
        string html = "<ol class=\"contains-task-list\"><li class=\"task-list-item\"><input class=\"task-list-item-checkbox\" type=\"checkbox\" disabled checked><p>Check</p><details open><summary>More</summary><blockquote class=\"callout note\"><p><strong>Watch</strong></p><p>Body</p>"
            + MarkdownVisualContract.BuildElementHtml(
                "canvas",
                "omd-visual omd-chart",
                MarkdownSemanticKinds.Chart,
                "ix-chart",
                MarkdownVisualContract.CreatePayload(payload))
            + "</blockquote></details></li></ol>";

        var markdownDocument = MarkdownReader.Parse(markdown, CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart));
        var htmlDocument = html.LoadFromHtml();

        var markdownList = Assert.IsType<OrderedListBlock>(Assert.Single(markdownDocument.Blocks));
        var htmlList = Assert.IsType<OrderedListBlock>(Assert.Single(htmlDocument.Blocks));

        var markdownItem = Assert.Single(markdownList.Items);
        var htmlItem = Assert.Single(htmlList.Items);

        Assert.Equal(markdownItem.Content.RenderMarkdown(), htmlItem.Content.RenderMarkdown());
        Assert.Equal(markdownItem.IsTask, htmlItem.IsTask);
        Assert.Equal(markdownItem.Checked, htmlItem.Checked);
        Assert.Equal(MarkdownAstParityFormatter.DescribeBlocks(markdownItem.Children), MarkdownAstParityFormatter.DescribeBlocks(htmlItem.Children));
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
        Assert.Equal(MarkdownAstParityFormatter.DescribeBlocks(markdownDetails.ChildBlocks), MarkdownAstParityFormatter.DescribeBlocks(htmlDetails.ChildBlocks));
    }

    [Fact]
    public void Semantic_Fenced_Block_Ast_Parity_Holds_Inside_Quotes() {
        const string payload = "{\r\n  \"type\": \"bar\",\r\n  \"data\": { \"labels\": [\"A\"], \"datasets\": [{ \"label\": \"Count\", \"data\": [1] }] }\r\n}";
        string quotedPayload = payload.Replace("\r\n", "\n").Replace("\n", "\n> ");
        string markdown = "> Intro\n>\n> ```ix-chart\n> "
            + quotedPayload
            + "\n> ```\n>\n> Tail\n";
        string html = "<blockquote><p>Intro</p>"
            + MarkdownVisualContract.BuildElementHtml(
                "canvas",
                "omd-visual omd-chart",
                MarkdownSemanticKinds.Chart,
                "ix-chart",
                MarkdownVisualContract.CreatePayload(payload))
            + "<p>Tail</p></blockquote>";

        var markdownDocument = MarkdownReader.Parse(markdown, CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart));
        var htmlDocument = html.LoadFromHtml();

        var markdownQuote = Assert.IsType<QuoteBlock>(Assert.Single(markdownDocument.Blocks));
        var htmlQuote = Assert.IsType<QuoteBlock>(Assert.Single(htmlDocument.Blocks));

        Assert.Equal(MarkdownAstParityFormatter.DescribeBlocks(markdownQuote.ChildBlocks), MarkdownAstParityFormatter.DescribeBlocks(htmlQuote.ChildBlocks));
    }

    [Fact]
    public void Semantic_Fenced_Block_Ast_Parity_Holds_Inside_Callouts() {
        const string payload = "{\r\n  \"type\": \"bar\",\r\n  \"data\": { \"labels\": [\"A\"], \"datasets\": [{ \"label\": \"Count\", \"data\": [1] }] }\r\n}";
        string quotedPayload = payload.Replace("\r\n", "\n").Replace("\n", "\n> ");
        string markdown = "> [!NOTE] Watch\n> Intro\n>\n> ```ix-chart\n> "
            + quotedPayload
            + "\n> ```\n>\n> Tail\n";
        string html = "<blockquote class=\"callout note\"><p><strong>Watch</strong></p><p>Intro</p>"
            + MarkdownVisualContract.BuildElementHtml(
                "canvas",
                "omd-visual omd-chart",
                MarkdownSemanticKinds.Chart,
                "ix-chart",
                MarkdownVisualContract.CreatePayload(payload))
            + "<p>Tail</p></blockquote>";

        var markdownDocument = MarkdownReader.Parse(markdown, CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart));
        var htmlDocument = html.LoadFromHtml();

        var markdownCallout = Assert.IsType<CalloutBlock>(Assert.Single(markdownDocument.Blocks));
        var htmlCallout = Assert.IsType<CalloutBlock>(Assert.Single(htmlDocument.Blocks));

        Assert.Equal(markdownCallout.Kind, htmlCallout.Kind);
        Assert.Equal(markdownCallout.TitleInlines.RenderMarkdown(), htmlCallout.TitleInlines.RenderMarkdown());
        Assert.Equal(MarkdownAstParityFormatter.DescribeBlocks(markdownCallout.ChildBlocks), MarkdownAstParityFormatter.DescribeBlocks(htmlCallout.ChildBlocks));
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
            MarkdownAstParityFormatter.DescribeBlocks(markdownTable.RowCells[0][0].Blocks),
            MarkdownAstParityFormatter.DescribeBlocks(htmlTable.RowCells[0][0].Blocks));
    }

    private static void AssertSemanticBlockParity(MarkdownDoc markdownDocument, MarkdownDoc htmlDocument) {
        var markdownBlock = Assert.IsType<SemanticFencedBlock>(Assert.Single(markdownDocument.Blocks));
        var htmlBlock = Assert.IsType<SemanticFencedBlock>(Assert.Single(htmlDocument.Blocks));

        Assert.Equal(markdownBlock.SemanticKind, htmlBlock.SemanticKind);
        Assert.Equal(markdownBlock.Language, htmlBlock.Language);
        Assert.Equal(markdownBlock.InfoString, htmlBlock.InfoString);
        Assert.Equal(markdownBlock.Content, htmlBlock.Content);
        Assert.Equal(markdownBlock.Caption, htmlBlock.Caption);
    }

    private static MarkdownReaderOptions CreateSemanticOptions(string language, string semanticKind) {
        var options = new MarkdownReaderOptions();
        options.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
            "Semantic AST",
            new[] { language },
            context => new SemanticFencedBlock(semanticKind, context.InfoString, context.Content, context.Caption)));
        return options;
    }
}
