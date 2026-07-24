namespace OfficeIMO.Markdown;

internal static class MarkdownListRendering {
    internal static string RenderMarkdown(
        MarkdownAttributeSet? attributes,
        IReadOnlyList<ListItem> items,
        Func<ListItem, int, string> markerFactory) {
        var sb = new System.Text.StringBuilder();
        var attributeBlock = MarkdownAttributeBlockRenderer.RenderInlineTrailing(attributes);
        if (!string.IsNullOrEmpty(attributeBlock)) {
            sb.Append(attributeBlock).AppendLine();
        }

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

            AppendItemMarkdown(sb, RenderItemMarkdown(item), baseIndent, firstPrefix, contPrefix);
            AppendChildrenMarkdown(sb, item, baseIndent, contPrefix);
        }

        return sb.ToString().TrimEnd();
    }

    internal static string RenderHtml(
        string listTag,
        IReadOnlyList<ListItem> items,
        MarkdownAttributeSet? attributes,
        Func<int, string> topLevelAttributesFactory,
        bool renderItemAttributes = false) {
        var sb = new System.Text.StringBuilder();
        renderItemAttributes = renderItemAttributes || HtmlRenderContext.RenderListItemAttributes;

        var scopeStarts = new List<int>();
        var taskScopes = new HashSet<(int StartIndex, int Level)>();
        var looseScopes = new HashSet<(int StartIndex, int Level)>();
        for (int index = 0; index < items.Count; index++) {
            int level = Math.Max(0, items[index].Level);
            if (scopeStarts.Count > level + 1) {
                scopeStarts.RemoveRange(level + 1, scopeStarts.Count - level - 1);
            }

            while (scopeStarts.Count <= level) {
                scopeStarts.Add(index);
            }

            var scopeKey = (scopeStarts[level], level);
            if (items[index].IsTask) {
                taskScopes.Add(scopeKey);
            }

            if (items[index].RequiresLooseListRendering()) {
                looseScopes.Add(scopeKey);
            }
        }

        bool ContainsTasksInScope(int startIndex, int level) => taskScopes.Contains((startIndex, level));

        bool IsLooseInScope(int startIndex, int level) => looseScopes.Contains((startIndex, level));

        void AppendOpenList(int startIndex, int level, bool isTopLevel) {
            var options = HtmlRenderContext.Options;
            var containsTasks = ContainsTasksInScope(startIndex, level);
            var useGitHubTaskListHtml = options?.GitHubTaskListHtml == true;
            sb.Append('<').Append(listTag);
            if (isTopLevel) {
                var taskClasses = containsTasks && !useGitHubTaskListHtml
                    ? new[] { "contains-task-list" }
                    : null;
                sb.Append(MarkdownHtmlAttributes.Render(attributes, options, additionalClasses: taskClasses));
                var attrs = topLevelAttributesFactory(startIndex);
                if (!string.IsNullOrEmpty(attrs)) {
                    sb.Append(attrs);
                }
            } else if (containsTasks && !useGitHubTaskListHtml) {
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
                if (liOpen) { sb.Append("</li>"); }
                for (int k = currentLevel; k > level; k--) { sb.Append("</").Append(listTag).Append("></li>"); }
                currentLevel = level;
            } else {
                if (liOpen) { sb.Append("</li>"); }
            }

            int scopeStart = scopeStartByLevel[level];
            bool renderLoose = IsLooseInScope(scopeStart, level);
            var options = HtmlRenderContext.Options;
            bool useGitHubTaskListHtml = options?.GitHubTaskListHtml == true;
            var itemClasses = item.IsTask && !useGitHubTaskListHtml
                ? new[] { "task-list-item" }
                : null;
            var itemAttributes = renderItemAttributes
                ? item.Attributes
                : null;
            sb.Append("<li")
                .Append(MarkdownHtmlAttributes.Render(itemAttributes, options, additionalClasses: itemClasses))
                .Append('>')
                .Append(item.RenderHtml(renderLoose, renderGenericAttributeConsumedWhitespace: !renderItemAttributes));
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

    private static string RenderItemMarkdown(ListItem item) {
        if (item == null) {
            return string.Empty;
        }

        var content = MarkdownEscaper.EscapeRenderedListItemLineStarts(item.RenderMarkdown());
        if (item.SyntaxChildren.Count == 0) {
            return content;
        }

        var definitions = new List<string>();
        for (int i = 0; i < item.SyntaxChildren.Count; i++) {
            var syntaxChild = item.SyntaxChildren[i];
            if (syntaxChild.Kind != MarkdownSyntaxKind.AbbreviationDefinition
                || string.IsNullOrWhiteSpace(syntaxChild.Literal)) {
                continue;
            }

            definitions.Add(syntaxChild.Literal!.TrimEnd('\r', '\n'));
        }

        if (definitions.Count == 0) {
            return content;
        }

        var prefix = string.Join("\n", definitions);
        return string.IsNullOrEmpty(content)
            ? prefix
            : prefix + "\n" + content;
    }

    private static void AppendChildrenMarkdown(System.Text.StringBuilder sb, ListItem item, string baseIndent, string contPrefix) {
        if (item.NestedBlocks.Count == 0) return;

        for (int c = 0; c < item.NestedBlocks.Count; c++) {
            var child = item.NestedBlocks[c];
            var childMd = MarkdownBlockRenderDispatcher.RenderMarkdown(child);
            if (string.IsNullOrWhiteSpace(childMd)) continue;

            var renderAsTightChildContinuation = ShouldRenderAsTightChildContinuation(item, child);
            if (!renderAsTightChildContinuation) {
                sb.Append(baseIndent).AppendLine();
            }

            var lines = childMd.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            for (int i = 0; i < lines.Length; i++) {
                var line = lines[i];
                if (line.Length == 0) sb.Append(baseIndent).AppendLine();
                else sb.Append(contPrefix).AppendLine(line);
            }

            if (!renderAsTightChildContinuation || c + 1 < item.NestedBlocks.Count) {
                sb.Append(baseIndent).AppendLine();
            }
        }
    }

    private static bool ShouldRenderAsTightChildContinuation(ListItem item, IMarkdownBlock child) =>
        (child is TableBlock || child is QuoteBlock || child is CustomContainerBlock) &&
        item != null &&
        !item.RequiresLooseListRendering() &&
        item.AdditionalParagraphs.Count == 0 &&
        item.Content.Nodes.Count > 0;
}
