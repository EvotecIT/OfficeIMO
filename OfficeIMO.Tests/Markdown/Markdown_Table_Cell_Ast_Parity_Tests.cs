using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.MarkdownRenderer;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_Table_Cell_Ast_Parity_Tests {
    [Fact]
    public void Table_Cell_Ast_Parity_Holds_For_Mixed_Block_Content() {
        const string markdown = """
| Section | Notes |
| --- | --- |
| Alpha | Intro<br><br>> Quoted<br><br>- first<br>- second |
""";
        const string html = "<table><tr><th>Section</th><th>Notes</th></tr><tr><td>Alpha</td><td><p>Intro</p><blockquote><p>Quoted</p></blockquote><ul><li>first</li><li>second</li></ul></td></tr></table>";

        AssertTableCellAstParity(markdown, html, rowIndex: 0, cellIndex: 1);
    }

    [Fact]
    public void Table_Cell_Ast_Parity_Holds_For_Code_Block_Content() {
        const string markdown = """
| Section | Notes |
| --- | --- |
| Alpha | Intro<br><br>```text<br>code line 1<br>code line 2<br>``` |
""";
        const string html = "<table><tr><th>Section</th><th>Notes</th></tr><tr><td>Alpha</td><td><p>Intro</p><pre><code class=\"language-text\">code line 1\ncode line 2</code></pre></td></tr></table>";

        AssertTableCellAstParity(markdown, html, rowIndex: 0, cellIndex: 1);
    }

    [Fact]
    public void Table_Cell_Ast_Parity_Holds_For_Details_Block_Content() {
        const string markdown = """
| Notes |
| --- |
| <details open><summary>More</summary>Alpha</details> |
""";
        const string html = "<table><tr><th>Notes</th></tr><tr><td><details open><summary>More</summary><p>Alpha</p></details></td></tr></table>";

        AssertTableCellAstParity(markdown, html, rowIndex: 0, cellIndex: 0);
    }

    [Fact]
    public void Table_Cell_Ast_Parity_Holds_For_Callout_Block_Content() {
        const string markdown = """
| Notes |
| --- |
| > [!WARNING] Watch<br>> Body |
""";
        const string html = "<table><tr><th>Notes</th></tr><tr><td><blockquote class=\"callout warning\"><p><strong>Watch</strong></p><p>Body</p></blockquote></td></tr></table>";

        AssertTableCellAstParity(markdown, html, rowIndex: 0, cellIndex: 0);
    }

    [Fact]
    public void Table_Cell_Ast_Parity_Holds_For_Quote_Followed_By_Callout_Content() {
        const string markdown = """
| Notes |
| --- |
| > Quoted<br><br>> [!WARNING] Watch<br>> Body |
""";
        const string html = "<table><tr><th>Notes</th></tr><tr><td><blockquote><p>Quoted</p></blockquote><blockquote class=\"callout warning\"><p><strong>Watch</strong></p><p>Body</p></blockquote></td></tr></table>";

        AssertTableCellAstParity(markdown, html, rowIndex: 0, cellIndex: 0);
    }

    [Fact]
    public void Table_Cell_Ast_Parity_Holds_For_Details_With_Quote_And_Callout_Content() {
        const string markdown = """
| Notes |
| --- |
| <details open><summary>More</summary><br><br>> Quoted<br><br>> [!WARNING] Watch<br>> Body<br></details> |
""";
        const string html = "<table><tr><th>Notes</th></tr><tr><td><details open><summary>More</summary><blockquote><p>Quoted</p></blockquote><blockquote class=\"callout warning\"><p><strong>Watch</strong></p><p>Body</p></blockquote></details></td></tr></table>";

        AssertTableCellAstParity(markdown, html, rowIndex: 0, cellIndex: 0);
    }

    [Fact]
    public void Table_Cell_Ast_Parity_Holds_For_Ordered_List_Content() {
        const string markdown = """
| Notes |
| --- |
| Intro<br><br>1. first<br>2. second |
""";
        const string html = "<table><tr><th>Notes</th></tr><tr><td><p>Intro</p><ol><li>first</li><li>second</li></ol></td></tr></table>";

        AssertTableCellAstParity(markdown, html, rowIndex: 0, cellIndex: 0);
    }

    [Fact]
    public void Table_Cell_Ast_Parity_Holds_For_Ordered_Task_List_With_Semantic_Block_Content() {
        const string payload = "{\r\n  \"type\": \"bar\",\r\n  \"data\": { \"labels\": [\"A\"], \"datasets\": [{ \"label\": \"Count\", \"data\": [1] }] }\r\n}";
        string cellPayload = payload.Replace("\r\n", "<br>   > ").Replace('\r', '\n').Replace("\n", "<br>   > ");
        string markdown = "| Notes |\n| --- |\n| 1. [x] Check<br><br>   <details open><br>   <summary>More</summary><br><br>   > [!NOTE] Watch<br>   > Body<br>   ><br>   > ```ix-chart<br>   > "
            + cellPayload
            + "<br>   > ```<br><br>   </details> |\n";
        string html = "<table><tr><th>Notes</th></tr><tr><td><ol class=\"contains-task-list\"><li class=\"task-list-item\"><input class=\"task-list-item-checkbox\" type=\"checkbox\" disabled checked><p>Check</p><details open><summary>More</summary><blockquote class=\"callout note\"><p><strong>Watch</strong></p><p>Body</p>"
            + MarkdownVisualContract.BuildElementHtml(
                "canvas",
                "omd-visual omd-chart",
                MarkdownSemanticKinds.Chart,
                "ix-chart",
                MarkdownVisualContract.CreatePayload(payload))
            + "</blockquote></details></li></ol></td></tr></table>";

        AssertTableCellAstParity(markdown, html, rowIndex: 0, cellIndex: 0, options: CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart));
    }

    [Fact]
    public void Table_Cell_Ast_Parity_Holds_For_Quote_Followed_By_Semantic_Block_Content() {
        const string payload = "{\r\n  \"type\": \"bar\",\r\n  \"data\": { \"labels\": [\"A\"], \"datasets\": [{ \"label\": \"Count\", \"data\": [1] }] }\r\n}";
        string cellPayload = payload.Replace("\r\n", "<br>").Replace('\r', '\n').Replace("\n", "<br>");
        string markdown = "| Notes |\n| --- |\n| > Quoted<br><br>```ix-chart<br>"
            + cellPayload
            + "<br>``` |\n";
        string html = "<table><tr><th>Notes</th></tr><tr><td><blockquote><p>Quoted</p></blockquote>"
            + MarkdownVisualContract.BuildElementHtml(
                "canvas",
                "omd-visual omd-chart",
                MarkdownSemanticKinds.Chart,
                "ix-chart",
                MarkdownVisualContract.CreatePayload(payload))
            + "</td></tr></table>";

        AssertTableCellAstParity(markdown, html, rowIndex: 0, cellIndex: 0, options: CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart));
    }

    [Fact]
    public void Table_Cell_Ast_Parity_Holds_For_Details_With_Callout_And_Semantic_Block_Content() {
        const string payload = "{\r\n  \"type\": \"bar\",\r\n  \"data\": { \"labels\": [\"A\"], \"datasets\": [{ \"label\": \"Count\", \"data\": [1] }] }\r\n}";
        string cellPayload = payload.Replace("\r\n", "<br>").Replace('\r', '\n').Replace("\n", "<br>");
        string markdown = "| Notes |\n| --- |\n| <details open><summary>More</summary><br><br>> [!WARNING] Watch<br>> Body<br><br>```ix-chart<br>"
            + cellPayload
            + "<br>```<br></details> |\n";
        string html = "<table><tr><th>Notes</th></tr><tr><td><details open><summary>More</summary>"
            + "<blockquote class=\"callout warning\"><p><strong>Watch</strong></p><p>Body</p></blockquote>"
            + MarkdownVisualContract.BuildElementHtml(
                "canvas",
                "omd-visual omd-chart",
                MarkdownSemanticKinds.Chart,
                "ix-chart",
                MarkdownVisualContract.CreatePayload(payload))
            + "</details></td></tr></table>";

        AssertTableCellAstParity(markdown, html, rowIndex: 0, cellIndex: 0, options: CreateSemanticOptions("ix-chart", MarkdownSemanticKinds.Chart));
    }

    private static void AssertTableCellAstParity(string markdown, string html, int rowIndex, int cellIndex, MarkdownReaderOptions? options = null) {
        var markdownDocument = MarkdownReader.Parse(markdown, options);
        var htmlDocument = html.LoadFromHtml();

        var markdownTable = Assert.IsType<TableBlock>(Assert.Single(markdownDocument.Blocks));
        var htmlTable = Assert.IsType<TableBlock>(Assert.Single(htmlDocument.Blocks));

        string markdownSummary = MarkdownAstParityFormatter.DescribeBlocks(markdownTable.RowCells[rowIndex][cellIndex].Blocks);
        string htmlSummary = MarkdownAstParityFormatter.DescribeBlocks(htmlTable.RowCells[rowIndex][cellIndex].Blocks);

        Assert.Equal(markdownSummary, htmlSummary);
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
