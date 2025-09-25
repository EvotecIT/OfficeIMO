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
        var sb = new System.Text.StringBuilder();
        string startAttr = Start != 1 ? " start=\"" + Start + "\"" : string.Empty;
        sb.Append("<ol").Append(startAttr).Append(">");
        int currentLevel = 0;
        bool liOpen = false;
        for (int idx = 0; idx < Items.Count; idx++) {
            var item = Items[idx];
            int level = item.Level;
            if (level > currentLevel) {
                // Open nested lists inside the current <li>
                for (int k = currentLevel; k < level; k++) sb.Append("<ol>");
                currentLevel = level;
                // Do not close parent <li> — nested list belongs inside it
            } else if (level < currentLevel) {
                // Close current item, unwind lists and their parent <li> at each level
                if (liOpen) { sb.Append("</li>"); liOpen = false; }
                for (int k = currentLevel; k > level; k--) { sb.Append("</ol>").Append("</li>"); }
                currentLevel = level;
            } else {
                // Same level: close previous <li>
                if (liOpen) { sb.Append("</li>"); liOpen = false; }
            }
            // Open new list item
            sb.Append("<li>").Append(item.RenderHtml());
            liOpen = true;
        }
        // Close the last open <li>
        if (liOpen) { sb.Append("</li>"); liOpen = false; }
        // Close any remaining nested lists and their parent <li> tags
        for (int k = currentLevel; k > 0; k--) { sb.Append("</ol>").Append("</li>"); }
        // Close top-level list
        sb.Append("</ol>");
        return sb.ToString();
    }
}
