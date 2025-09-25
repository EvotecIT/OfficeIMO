namespace OfficeIMO.Markdown;

/// <summary>
/// Unordered list supporting plain items and task (checklist) items.
/// </summary>
public sealed class UnorderedListBlock : IMarkdownBlock {
    /// <summary>List items.</summary>
    public List<ListItem> Items { get; } = new List<ListItem>();
    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() => string.Join("\n", Items.Select(i => i.ToMarkdownListLine()));
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        var sb = new System.Text.StringBuilder();
        sb.Append("<ul>");
        int currentLevel = 0;
        bool liOpen = false;
        for (int idx = 0; idx < Items.Count; idx++) {
            var item = Items[idx];
            int level = item.Level;
            if (level > currentLevel) {
                for (int k = currentLevel; k < level; k++) sb.Append("<ul>");
                currentLevel = level;
                // keep parent <li> open
            } else if (level < currentLevel) {
                if (liOpen) { sb.Append("</li>"); liOpen = false; }
                for (int k = currentLevel; k > level; k--) { sb.Append("</ul>").Append("</li>"); }
                currentLevel = level;
            } else {
                if (liOpen) { sb.Append("</li>"); liOpen = false; }
            }
            // Compose content with optional task checkbox
            string content = item.RenderHtml();
            if (item.IsTask) {
                content = "<input type=\"checkbox\" disabled" + (item.Checked ? " checked" : string.Empty) + "> " + content;
            }
            sb.Append("<li>").Append(content);
            liOpen = true;
        }
        if (liOpen) { sb.Append("</li>"); liOpen = false; }
        for (int k = currentLevel; k > 0; k--) { sb.Append("</ul>").Append("</li>"); }
        sb.Append("</ul>");
        return sb.ToString();
    }
}
