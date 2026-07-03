using System.Globalization;
using System.Linq;
using System.Text;
using OfficeIMO.Markdown;

namespace OfficeIMO.Tests.MarkdownSuite;

internal static class MarkdownAstParityFormatter {
    public static string DescribeBlocks(IReadOnlyList<IMarkdownBlock> blocks) {
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
            case TableBlock table:
                AppendTable(sb, table, indent + 1);
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

    private static void AppendTable(StringBuilder sb, TableBlock table, int indent) {
        string prefix = new string(' ', indent * 2);
        if (table.HeaderCells.Count > 0) {
            sb.Append(prefix)
                .Append("headers: ")
                .AppendLine(string.Join(" || ", table.HeaderCells.Select(cell => "\"" + EscapeSingleLine(cell.Markdown) + "\"")));
        }

        for (int rowIndex = 0; rowIndex < table.RowCells.Count; rowIndex++) {
            sb.Append(prefix)
                .Append("row[")
                .Append(rowIndex.ToString(CultureInfo.InvariantCulture))
                .Append("]: ")
                .AppendLine(string.Join(" || ", table.RowCells[rowIndex].Select(cell => "\"" + EscapeSingleLine(cell.Markdown) + "\"")));
        }
    }

    private static string DescribeBlock(IMarkdownBlock block) {
        return block switch {
            HeadingBlock heading => $"Heading(level={heading.Level}, text=\"{EscapeSingleLine(heading.Text)}\")",
            ParagraphBlock paragraph => $"Paragraph(\"{EscapeSingleLine(paragraph.Inlines.RenderMarkdown())}\")",
            CalloutBlock callout => $"Callout(kind={callout.Kind}, title=\"{EscapeSingleLine(callout.TitleInlines.RenderMarkdown())}\")",
            QuoteBlock => "Quote",
            UnorderedListBlock unordered => $"UnorderedList(items={unordered.Items.Count.ToString(CultureInfo.InvariantCulture)})",
            OrderedListBlock ordered => $"OrderedList(start={ordered.Start.ToString(CultureInfo.InvariantCulture)}, items={ordered.Items.Count.ToString(CultureInfo.InvariantCulture)})",
            CodeBlock code => $"Code(language={code.Language}, text=\"{EscapeSingleLine(code.Content)}\")",
            TableBlock table => $"Table(headers={table.HeaderCells.Count.ToString(CultureInfo.InvariantCulture)}, rows={table.RowCells.Count.ToString(CultureInfo.InvariantCulture)})",
            DefinitionListBlock definitionList => $"DefinitionList(entries={definitionList.Entries.Count.ToString(CultureInfo.InvariantCulture)})",
            FootnoteDefinitionBlock footnote => $"Footnote(label={footnote.Label})",
            DetailsBlock details => $"Details(open={details.Open.ToString().ToLowerInvariant()})",
            SemanticFencedBlock semantic => $"Semantic(kind={semantic.SemanticKind}, language={semantic.Language}, text=\"{EscapeSingleLine(AbbreviateSemanticContent(semantic.Content))}\")",
            _ => block.GetType().Name
        };
    }

    private static string EscapeSingleLine(string? value) {
        if (string.IsNullOrEmpty(value)) {
            return string.Empty;
        }

        return value!
            .Replace("\r\n", "\\n")
            .Replace('\r', '\n')
            .Replace("\n", "\\n");
    }

    private static string AbbreviateSemanticContent(string? value) {
        const int maxLength = 80;
        if (string.IsNullOrEmpty(value) || value!.Length <= maxLength) {
            return value ?? string.Empty;
        }

        return value.Substring(0, maxLength - 3) + "...";
    }
}
