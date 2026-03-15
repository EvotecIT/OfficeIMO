using System.Globalization;
using System.Text;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
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

    private static void AssertTableCellAstParity(string markdown, string html, int rowIndex, int cellIndex) {
        var markdownDocument = MarkdownReader.Parse(markdown);
        var htmlDocument = html.LoadFromHtml();

        var markdownTable = Assert.IsType<TableBlock>(Assert.Single(markdownDocument.Blocks));
        var htmlTable = Assert.IsType<TableBlock>(Assert.Single(htmlDocument.Blocks));

        string markdownSummary = DescribeBlocks(markdownTable.RowCells[rowIndex][cellIndex].Blocks);
        string htmlSummary = DescribeBlocks(htmlTable.RowCells[rowIndex][cellIndex].Blocks);

        Assert.Equal(markdownSummary, htmlSummary);
    }

    private static string DescribeBlocks(IReadOnlyList<IMarkdownBlock> blocks) {
        var sb = new StringBuilder();
        AppendBlocks(sb, blocks, 0);
        return sb.ToString().TrimEnd();
    }

    private static void AppendBlocks(StringBuilder sb, IReadOnlyList<IMarkdownBlock> blocks, int indent) {
        for (int i = 0; i < blocks.Count; i++) {
            AppendBlock(sb, blocks[i], indent, i);
        }
    }

    private static void AppendBlock(StringBuilder sb, IMarkdownBlock block, int indent, int index) {
        string prefix = new string(' ', indent * 2);
        sb.Append(prefix)
            .Append(index.ToString(CultureInfo.InvariantCulture))
            .Append(": ")
            .AppendLine(DescribeBlock(block));

        switch (block) {
            case QuoteBlock quote:
                AppendBlocks(sb, quote.ChildBlocks, indent + 1);
                break;
            case CalloutBlock callout:
                AppendBlocks(sb, callout.ChildBlocks, indent + 1);
                break;
            case DetailsBlock details:
                if (details.Summary != null) {
                    sb.Append(new string(' ', (indent + 1) * 2))
                        .Append("summary: ")
                        .AppendLine(EscapeSingleLine(details.Summary.Inlines.RenderMarkdown()));
                }
                AppendBlocks(sb, details.ChildBlocks, indent + 1);
                break;
            case UnorderedListBlock unordered:
                AppendListItems(sb, unordered.Items, indent + 1);
                break;
            case OrderedListBlock ordered:
                AppendListItems(sb, ordered.Items, indent + 1);
                break;
        }
    }

    private static void AppendListItems(StringBuilder sb, IReadOnlyList<ListItem> items, int indent) {
        string prefix = new string(' ', indent * 2);
        for (int i = 0; i < items.Count; i++) {
            var item = items[i];
            sb.Append(prefix)
                .Append("item[")
                .Append(i.ToString(CultureInfo.InvariantCulture))
                .Append("]: task=")
                .Append(item.IsTask ? (item.Checked ? "checked" : "unchecked") : "no")
                .Append(" content=\"")
                .Append(EscapeSingleLine(item.Content.RenderMarkdown()))
                .AppendLine("\"");

            AppendBlocks(sb, item.Children, indent + 1);
        }
    }

    private static string DescribeBlock(IMarkdownBlock block) {
        return block switch {
            ParagraphBlock paragraph => $"Paragraph(\"{EscapeSingleLine(paragraph.Inlines.RenderMarkdown())}\")",
            QuoteBlock => "Quote",
            UnorderedListBlock unordered => $"UnorderedList(items={unordered.Items.Count.ToString(CultureInfo.InvariantCulture)})",
            OrderedListBlock ordered => $"OrderedList(start={ordered.Start.ToString(CultureInfo.InvariantCulture)}, items={ordered.Items.Count.ToString(CultureInfo.InvariantCulture)})",
            CodeBlock code => $"Code(language={code.Language}, text=\"{EscapeSingleLine(code.Content)}\")",
            HeadingBlock heading => $"Heading(level={heading.Level}, text=\"{EscapeSingleLine(heading.Text)}\")",
            DetailsBlock details => $"Details(open={details.Open.ToString().ToLowerInvariant()})",
            CalloutBlock callout => $"Callout(kind={callout.Kind}, title=\"{EscapeSingleLine(callout.TitleInlines.RenderMarkdown())}\")",
            _ => block.GetType().Name
        };
    }

    private static string EscapeSingleLine(string? value) {
        return (value ?? string.Empty)
            .Replace("\\", "\\\\")
            .Replace("\r", "\\r")
            .Replace("\n", "\\n")
            .Replace("\"", "\\\"");
    }
}
