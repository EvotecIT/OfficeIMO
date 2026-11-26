namespace OfficeIMO.Markdown;

internal static class MarkdownListRendering {
    private static readonly string[] ContainsTaskListClasses = { "contains-task-list" };
    private static readonly string[] TaskListItemClasses = { "task-list-item" };

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
        string topLevelAttributes,
        bool renderItemAttributes = false) {
        var sb = new System.Text.StringBuilder();
        AppendHtml(sb, listTag, items, attributes, topLevelAttributes, renderItemAttributes);
        return sb.ToString();
    }

    internal static void AppendHtml(
        System.Text.StringBuilder sb,
        string listTag,
        IReadOnlyList<ListItem> items,
        MarkdownAttributeSet? attributes,
        string topLevelAttributes,
        bool renderItemAttributes = false) {
        renderItemAttributes = renderItemAttributes || HtmlRenderContext.RenderListItemAttributes;

        // The item index where the open list at each level started. A flat list has one scope, which
        // starts at the first item, so it needs no array.
        int maxLevel = 0;
        for (int index = 0; index < items.Count; index++) {
            maxLevel = Math.Max(maxLevel, items[index].Level);
        }

        int[]? scopeStarts = maxLevel == 0 ? null : new int[maxLevel + 1];
        int scopeCount = 0;
        // Most lists have neither task items nor loose items, so the sets are made on first use.
        HashSet<(int StartIndex, int Level)>? taskScopes = null;
        HashSet<(int StartIndex, int Level)>? looseScopes = null;
        for (int index = 0; index < items.Count; index++) {
            int level = Math.Max(0, items[index].Level);
            if (scopeCount > level + 1) {
                scopeCount = level + 1;
            }

            while (scopeCount <= level) {
                if (scopeStarts != null) {
                    scopeStarts[scopeCount] = index;
                }

                scopeCount++;
            }

            var scopeKey = (scopeStarts?[level] ?? 0, level);
            if (items[index].IsTask) {
                (taskScopes ??= new HashSet<(int StartIndex, int Level)>()).Add(scopeKey);
            }

            if (items[index].RequiresLooseListRendering()) {
                (looseScopes ??= new HashSet<(int StartIndex, int Level)>()).Add(scopeKey);
            }
        }

        bool ContainsTasksInScope(int startIndex, int level) => taskScopes?.Contains((startIndex, level)) == true;

        bool IsLooseInScope(int startIndex, int level) => looseScopes?.Contains((startIndex, level)) == true;

        void AppendOpenList(int startIndex, int level, bool isTopLevel) {
            var options = HtmlRenderContext.Options;
            var containsTasks = ContainsTasksInScope(startIndex, level);
            var useGitHubTaskListHtml = options?.GitHubTaskListHtml == true;
            sb.Append('<').Append(listTag);
            if (isTopLevel) {
                var taskClasses = containsTasks && !useGitHubTaskListHtml
                    ? ContainsTaskListClasses
                    : null;
                sb.Append(MarkdownHtmlAttributes.Render(attributes, options, additionalClasses: taskClasses));
                if (!string.IsNullOrEmpty(topLevelAttributes)) {
                    sb.Append(topLevelAttributes);
                }
            } else if (containsTasks && !useGitHubTaskListHtml) {
                sb.Append(" class=\"contains-task-list\"");
            }
            sb.Append('>');
        }

        AppendOpenList(0, 0, isTopLevel: true);
        var renderOptions = HtmlRenderContext.Options;
        bool needsInlineTaskMarker = renderOptions == null
            || renderOptions.CssDelivery == MarkdownCssDelivery.None
            || HtmlRenderer.GetEffectiveStyle(renderOptions) == MarkdownHtmlStyle.Plain;
        int[]? scopeStartByLevel = maxLevel == 0 ? null : new int[maxLevel + 1];
        int currentLevel = 0;
        bool liOpen = false;

        for (int idx = 0; idx < items.Count; idx++) {
            var item = items[idx];
            int level = item.Level;
            if (level > currentLevel) {
                for (int k = currentLevel + 1; k <= level; k++) {
                    AppendOpenList(idx, k, isTopLevel: false);
                    scopeStartByLevel![k] = idx;
                }
                currentLevel = level;
            } else if (level < currentLevel) {
                if (liOpen) { sb.Append("</li>"); }
                for (int k = currentLevel; k > level; k--) { sb.Append("</").Append(listTag).Append("></li>"); }
                currentLevel = level;
            } else {
                if (liOpen) { sb.Append("</li>"); }
            }

            int scopeStart = scopeStartByLevel?[level] ?? 0;
            bool renderLoose = IsLooseInScope(scopeStart, level);
            var options = HtmlRenderContext.Options;
            bool useGitHubTaskListHtml = options?.GitHubTaskListHtml == true;
            var itemClasses = item.IsTask && !useGitHubTaskListHtml
                ? TaskListItemClasses
                : null;
            var itemAttributes = renderItemAttributes
                ? item.Attributes
                : null;
            string taskMarkerStyle = string.Empty;
            if (item.IsTask && !useGitHubTaskListHtml && needsInlineTaskMarker) {
                string? authoredStyle = itemAttributes?.GetAttribute("style");
                if (itemAttributes != null && itemAttributes.TryGetAttribute("style", out _)) {
                    itemAttributes = MarkdownAttributeSet.Create(itemAttributes.ElementId, itemAttributes.Classes,
                        itemAttributes.Attributes.Where(pair => !string.Equals(pair.Key, "style", StringComparison.OrdinalIgnoreCase)));
                }
                // The checkbox is the task marker, including without a presentation stylesheet.
                // Retain authored CSS after the safe default so source declarations win.
                string retainedStyle = (options?.RawHtmlHandling ?? MarkdownRawHtmlHandling.Allow) == MarkdownRawHtmlHandling.Allow
                    ? authoredStyle ?? string.Empty
                    : string.Empty;
                taskMarkerStyle = " style=\"list-style:none;" + HtmlTextEncoder.Encode(retainedStyle, options) + "\"";
            }
            sb.Append("<li")
                .Append(MarkdownHtmlAttributes.Render(itemAttributes, options, additionalClasses: itemClasses))
                .Append(taskMarkerStyle)
                .Append('>');
            item.AppendHtml(sb, renderLoose, renderGenericAttributeConsumedWhitespace: !renderItemAttributes);
            liOpen = true;
        }

        if (liOpen) { sb.Append("</li>"); }
        for (int k = currentLevel; k > 0; k--) { sb.Append("</").Append(listTag).Append("></li>"); }
        sb.Append("</").Append(listTag).Append('>');
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

        var content = EscapeItemLineStarts(item);
        var syntaxChildren = item.SyntaxChildrenForWriting;
        if (syntaxChildren == null || syntaxChildren.Count == 0) {
            return content;
        }

        var definitions = new List<string>();
        for (int i = 0; i < syntaxChildren.Count; i++) {
            var syntaxChild = syntaxChildren[i];
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

    private static string EscapeItemLineStarts(ListItem item) {
        string markdown = item.RenderMarkdown();
        if (item.Content.Nodes.Count == 0 || item.Content.Nodes[0] is not HtmlRawInline anchor ||
            !anchor.Html.StartsWith("<a ", StringComparison.OrdinalIgnoreCase) ||
            !anchor.Html.EndsWith("</a>", StringComparison.OrdinalIgnoreCase) ||
            anchor.Html.IndexOf('\r') >= 0 || anchor.Html.IndexOf('\n') >= 0 ||
            !markdown.StartsWith(anchor.Html, StringComparison.Ordinal))
            return MarkdownEscaper.EscapeRenderedListItemLineStarts(markdown);

        // A complete inline anchor stays active. Continue escaping the other lines as usual.
        int newline = markdown.IndexOf('\n');
        return newline < 0 ? markdown : markdown.Substring(0, newline + 1) +
            MarkdownEscaper.EscapeRenderedListItemLineStarts(markdown.Substring(newline + 1));
    }

    private static void AppendChildrenMarkdown(System.Text.StringBuilder sb, ListItem item, string baseIndent, string contPrefix) {
        if (item.NestedBlockCount == 0) return;

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
        !item.HasAdditionalParagraphs &&
        item.Content.Nodes.Count > 0;
}
