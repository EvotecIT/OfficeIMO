namespace OfficeIMO.Markdown;

internal static class MarkdownListRendering {
    internal static string RenderMarkdown(
        IReadOnlyList<ListItem> items,
        Func<ListItem, int, string> markerFactory) {
        var sb = new System.Text.StringBuilder();
        int topLevelIndex = 0;

        for (int idx = 0; idx < items.Count; idx++) {
            var item = items[idx];
            string baseIndent = new string(' ', item.Level * 2);
            string marker = markerFactory(item, topLevelIndex);
            if (item.Level == 0) {
                topLevelIndex++;
            }

            string firstPrefix = baseIndent + marker;
            string contPrefix = baseIndent + new string(' ', marker.Length);

            AppendItemMarkdown(sb, item.RenderMarkdown(), baseIndent, firstPrefix, contPrefix);
            AppendChildrenMarkdown(sb, item, baseIndent, contPrefix);
        }

        return sb.ToString().TrimEnd();
    }

    internal static string RenderHtml(
        string listTag,
        IReadOnlyList<ListItem> items,
        Func<int, string> topLevelAttributesFactory) {
        var sb = new System.Text.StringBuilder();

        bool ContainsTasksInScope(int startIndex, int level) {
            for (int i = startIndex; i < items.Count; i++) {
                var it = items[i];
                if (it.Level < level) break;
                if (it.Level == level && it.IsTask) return true;
            }
            return false;
        }

        bool IsLooseInScope(int startIndex, int level) {
            for (int i = startIndex; i < items.Count; i++) {
                var it = items[i];
                if (it.Level < level) break;
                if (it.Level != level) continue;
                if (it.RequiresLooseListRendering()) return true;
            }
            return false;
        }

        void AppendOpenList(int startIndex, int level, bool isTopLevel) {
            sb.Append('<').Append(listTag);
            if (isTopLevel) {
                var attrs = topLevelAttributesFactory(startIndex);
                if (!string.IsNullOrEmpty(attrs)) {
                    sb.Append(attrs);
                }
            }
            if (ContainsTasksInScope(startIndex, level)) {
                sb.Append(" class=\"contains-task-list\"");
            }
            sb.Append('>');
        }

        AppendOpenList(0, 0, isTopLevel: true);
        var scopeStartByLevel = new List<int> { 0 };
        int currentLevel = 0;
        bool liOpen = false;

        for (int idx = 0; idx < items.Count; idx++) {
            var item = items[idx];
            int level = item.Level;
            if (level > currentLevel) {
                for (int k = currentLevel + 1; k <= level; k++) {
                    AppendOpenList(idx, k, isTopLevel: false);
                    if (scopeStartByLevel.Count <= k) scopeStartByLevel.Add(idx);
                    else scopeStartByLevel[k] = idx;
                }
                currentLevel = level;
            } else if (level < currentLevel) {
                if (liOpen) { sb.Append("</li>"); liOpen = false; }
                for (int k = currentLevel; k > level; k--) { sb.Append("</").Append(listTag).Append("></li>"); }
                currentLevel = level;
            } else {
                if (liOpen) { sb.Append("</li>"); liOpen = false; }
            }

            int scopeStart = scopeStartByLevel[level];
            bool renderLoose = IsLooseInScope(scopeStart, level);
            sb.Append(item.IsTask ? "<li class=\"task-list-item\">" : "<li>").Append(item.RenderHtml(renderLoose));
            liOpen = true;
        }

        if (liOpen) { sb.Append("</li>"); }
        for (int k = currentLevel; k > 0; k--) { sb.Append("</").Append(listTag).Append("></li>"); }
        sb.Append("</").Append(listTag).Append('>');
        return sb.ToString();
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
}
