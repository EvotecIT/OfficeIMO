using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Markdown;

/// <summary>
/// Ordered (numbered) list.
/// </summary>
public sealed class OrderedListBlock : IMarkdownBlock {
    /// <summary>Items within the ordered list.</summary>
    public List<ListItem> Items { get; } = new List<ListItem>();
    /// <summary>Starting number (default 1).</summary>
    public int Start { get; set; } = 1;

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        int i = Start;
        var sb = new System.Text.StringBuilder();
        foreach (var item in Items) {
            var indent = new string(' ', item.Level * 2);
            if (item.Level == 0) {
                sb.Append(indent).Append(i++).Append(". ").Append(item.RenderMarkdown()).Append('\n');
            } else {
                // For nested levels, emit "1." which most renderers normalize visually
                sb.Append(indent).Append("1. ").Append(item.RenderMarkdown()).Append('\n');
            }
        }
        return sb.ToString().TrimEnd();
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        string startAttr = Start != 1 ? " start=\"" + Start + "\"" : string.Empty;
        return "<ol" + startAttr + ">" + string.Concat(Items.Select(i => "<li>" + i.RenderHtml() + "</li>")) + "</ol>";
    }
}
