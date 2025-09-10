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
        return string.Join("\n", Items.Select(item => (i++) + ". " + item.RenderMarkdown()));
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        string startAttr = Start != 1 ? " start=\"" + Start + "\"" : string.Empty;
        return "<ol" + startAttr + ">" + string.Concat(Items.Select(i => "<li>" + i.RenderHtml() + "</li>")) + "</ol>";
    }
}
