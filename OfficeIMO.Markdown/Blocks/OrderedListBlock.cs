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
            string baseMarker = item.Level == 0 ? (i++).ToString(System.Globalization.CultureInfo.InvariantCulture) + ". " : "1. ";
            string marker = item.IsTask
                ? baseMarker + "[" + (item.Checked ? "x" : " ") + "] "
                : baseMarker;
            string firstPrefix = indent + marker;
            string contPrefix = indent + new string(' ', marker.Length);

            AppendItemMarkdown(sb, item.RenderMarkdown(), indent, firstPrefix, contPrefix);
            AppendChildrenMarkdown(sb, item, indent, contPrefix);
        }
        return sb.ToString().TrimEnd();
    }

    private static void AppendItemMarkdown(System.Text.StringBuilder sb, string content, string baseIndent, string firstPrefix, string contPrefix) {
        if (content == null) content = string.Empty;
        var lines = content.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        for (int li = 0; li < lines.Length; li++) {
            var line = lines[li];
            if (li == 0) {
                sb.Append(firstPrefix).AppendLine(line);
                continue;
            }
            if (line.Length == 0) {
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
        string startAttr = Start != 1 ? " start=\"" + Start + "\"" : string.Empty;

        bool ContainsTasksInScope(int startIndex, int level) {
            for (int i = startIndex; i < Items.Count; i++) {
                var it = Items[i];
                if (it.Level < level) break;
                if (it.Level == level && it.IsTask) return true;
            }
            return false;
        }

        void AppendOpenOl(int startIndex, int level, bool isTopLevel) {
            var cls = ContainsTasksInScope(startIndex, level) ? " class=\"contains-task-list\"" : string.Empty;
            if (isTopLevel) sb.Append("<ol").Append(startAttr).Append(cls).Append(">");
            else sb.Append("<ol").Append(cls).Append(">");
        }

        AppendOpenOl(0, 0, isTopLevel: true);
        int currentLevel = 0;
        bool liOpen = false;
        for (int idx = 0; idx < Items.Count; idx++) {
            var item = Items[idx];
            int level = item.Level;
            if (level > currentLevel) {
                // Open nested lists inside the current <li>
                for (int k = currentLevel + 1; k <= level; k++) AppendOpenOl(idx, k, isTopLevel: false);
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
            sb.Append(item.IsTask ? "<li class=\"task-list-item\">" : "<li>").Append(item.RenderHtml());
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
