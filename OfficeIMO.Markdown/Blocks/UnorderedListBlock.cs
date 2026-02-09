namespace OfficeIMO.Markdown;

/// <summary>
/// Unordered list supporting plain items and task (checklist) items.
/// </summary>
public sealed class UnorderedListBlock : IMarkdownBlock {
    /// <summary>List items.</summary>
    public List<ListItem> Items { get; } = new List<ListItem>();
    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        var sb = new System.Text.StringBuilder();
        for (int idx = 0; idx < Items.Count; idx++) {
            var item = Items[idx];
            string baseIndent = new string(' ', item.Level * 2);

            string marker = item.IsTask
                ? "- [" + (item.Checked ? "x" : " ") + "] "
                : "- ";

            string firstPrefix = baseIndent + marker;
            string contPrefix = baseIndent + new string(' ', marker.Length);

            AppendItemMarkdown(sb, item.RenderMarkdown(), baseIndent, firstPrefix, contPrefix);
            AppendChildrenMarkdown(sb, item, baseIndent, contPrefix);
        }
        return sb.ToString().TrimEnd();
    }

    private static void AppendItemMarkdown(System.Text.StringBuilder sb, string content, string baseIndent, string firstPrefix, string contPrefix) {
        if (content == null) content = string.Empty;
        var lines = content.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        for (int i = 0; i < lines.Length; i++) {
            string line = lines[i];
            if (i == 0) {
                sb.Append(firstPrefix).AppendLine(line);
                continue;
            }
            if (line.Length == 0) {
                // Indent blank lines so they remain part of the list item.
                sb.Append(baseIndent).AppendLine();
            } else {
                sb.Append(contPrefix).AppendLine(line);
            }
        }
    }

    private static void AppendChildrenMarkdown(System.Text.StringBuilder sb, ListItem item, string baseIndent, string contPrefix) {
        if (item.Children.Count == 0) return;

        for (int c = 0; c < item.Children.Count; c++) {
            var child = item.Children[c];
            var childMd = child.RenderMarkdown();
            if (string.IsNullOrWhiteSpace(childMd)) continue;

            // Separate paragraph content from child blocks.
            sb.Append(baseIndent).AppendLine();

            var lines = childMd.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            for (int i = 0; i < lines.Length; i++) {
                var line = lines[i];
                if (line.Length == 0) sb.Append(baseIndent).AppendLine();
                else sb.Append(contPrefix).AppendLine(line);
            }
        }
    }
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
            sb.Append("<li>").Append(item.RenderHtml());
            liOpen = true;
        }
        if (liOpen) { sb.Append("</li>"); liOpen = false; }
        for (int k = currentLevel; k > 0; k--) { sb.Append("</ul>").Append("</li>"); }
        sb.Append("</ul>");
        return sb.ToString();
    }
}
