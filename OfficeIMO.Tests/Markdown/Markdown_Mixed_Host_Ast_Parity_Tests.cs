using System.Globalization;
using System.Text;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
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

    private static void AssertDocumentAstParity(string markdown, string html) {
        var markdownDocument = MarkdownReader.Parse(markdown);
        var htmlDocument = html.LoadFromHtml();

        Assert.Equal(
            DescribeBlocks(markdownDocument.Blocks),
            DescribeBlocks(htmlDocument.Blocks));
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
            SemanticFencedBlock semantic => $"Semantic(kind={semantic.SemanticKind}, language={semantic.Language}, text=\"{EscapeSingleLine(semantic.Content)}\")",
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
