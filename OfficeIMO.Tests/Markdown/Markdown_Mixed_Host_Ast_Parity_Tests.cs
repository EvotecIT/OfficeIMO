using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.MarkdownRenderer;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_Mixed_Host_Ast_Parity_Tests {
    [Fact]
    public void Mixed_Host_Ast_Parity_Holds_For_Quote_With_Details_And_List() {
        const string markdown = """
> Intro
>
> <details open>
> <summary>More</summary>
>
> Hidden
> </details>
>
> - first
> - second
""";
        const string html = "<blockquote><p>Intro</p><details open><summary>More</summary><p>Hidden</p></details><ul><li>first</li><li>second</li></ul></blockquote>";

        AssertDocumentAstParity(markdown, html);
    }

    [Fact]
    public void Mixed_Host_Ast_Parity_Holds_For_Callout_With_Details_And_List() {
        const string markdown = """
> [!NOTE] Watch
> Intro
>
> <details open>
> <summary>More</summary>
>
> Hidden
> </details>
>
> - first
> - second
""";
        const string html = "<blockquote class=\"callout note\"><p><strong>Watch</strong></p><p>Intro</p><details open><summary>More</summary><p>Hidden</p></details><ul><li>first</li><li>second</li></ul></blockquote>";

        AssertDocumentAstParity(markdown, html);
    }

    [Fact]
    public void Mixed_Host_Ast_Parity_Holds_For_Details_With_Quote_And_Callout() {
        const string markdown = """
<details open>
<summary>More</summary>

> Quoted

> [!WARNING] Watch
> Body
</details>
""";
        const string html = "<details open><summary>More</summary><blockquote><p>Quoted</p></blockquote><blockquote class=\"callout warning\"><p><strong>Watch</strong></p><p>Body</p></blockquote></details>";

        AssertDocumentAstParity(markdown, html);
    }

    [Fact]
    public void Mixed_Host_Ast_Parity_Holds_For_Quote_With_Details_And_Semantic_Block() {
        const string payload = "{\r\n  \"type\": \"bar\",\r\n  \"data\": { \"labels\": [\"A\"], \"datasets\": [{ \"label\": \"Count\", \"data\": [1] }] }\r\n}";
        string quotedPayload = payload.Replace("\r\n", "\n").Replace("\n", "\n> ");
        string markdown = "> Intro\n>\n> <details open>\n> <summary>More</summary>\n>\n> Hidden\n>\n> ```ix-chart\n> "
            + quotedPayload
            + "\n> ```\n> </details>\n";
        string html = "<blockquote><p>Intro</p><details open><summary>More</summary><p>Hidden</p>"
            + MarkdownVisualContract.BuildElementHtml(
                "canvas",
                "omd-visual omd-chart",
                MarkdownSemanticKinds.Chart,
                "ix-chart",
                MarkdownVisualContract.CreatePayload(payload))
            + "</details></blockquote>";

        AssertDocumentAstParity(markdown, html, CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart));
    }

    [Fact]
    public void Mixed_Host_Ast_Parity_Holds_For_Callout_With_Details_And_Semantic_Block() {
        const string payload = "{\r\n  \"type\": \"bar\",\r\n  \"data\": { \"labels\": [\"A\"], \"datasets\": [{ \"label\": \"Count\", \"data\": [1] }] }\r\n}";
        string quotedPayload = payload.Replace("\r\n", "\n").Replace("\n", "\n> ");
        string markdown = "> [!NOTE] Watch\n> Intro\n>\n> <details open>\n> <summary>More</summary>\n>\n> Hidden\n>\n> ```ix-chart\n> "
            + quotedPayload
            + "\n> ```\n> </details>\n";
        string html = "<blockquote class=\"callout note\"><p><strong>Watch</strong></p><p>Intro</p><details open><summary>More</summary><p>Hidden</p>"
            + MarkdownVisualContract.BuildElementHtml(
                "canvas",
                "omd-visual omd-chart",
                MarkdownSemanticKinds.Chart,
                "ix-chart",
                MarkdownVisualContract.CreatePayload(payload))
            + "</details></blockquote>";

        AssertDocumentAstParity(markdown, html, CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart));
    }

    private static void AssertDocumentAstParity(string markdown, string html, MarkdownReaderOptions? options = null) {
        var markdownDocument = MarkdownReader.Parse(markdown, options);
        var htmlDocument = html.LoadFromHtml();

        Assert.Equal(
            MarkdownAstParityFormatter.DescribeBlocks(markdownDocument.Blocks),
            MarkdownAstParityFormatter.DescribeBlocks(htmlDocument.Blocks));
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
