namespace OfficeIMO.Markdown;

/// <summary>
/// Built-in block render extension registrations for portable or host-specific output shaping.
/// </summary>
public static class MarkdownBlockRenderBuiltInExtensions {
    /// <summary>Stable registration name for the portable callout markdown fallback.</summary>
    public const string PortableCalloutMarkdownName = "Portable.Callout.Markdown";

    /// <summary>Stable registration name for the portable callout HTML fallback.</summary>
    public const string PortableCalloutHtmlName = "Portable.Callout.Html";
    /// <summary>Stable registration name for the Markdig/GitHub-style alert HTML fallback.</summary>
    public const string MarkdigAlertHtmlName = "Markdig.Alert.Html";
    /// <summary>Stable registration name for the portable TOC HTML fallback.</summary>
    public const string PortableTocHtmlName = "Portable.Toc.Html";
    /// <summary>Stable registration name for the portable footnote section HTML fallback.</summary>
    public const string PortableFootnoteSectionHtmlName = "Portable.Footnotes.Html";
    /// <summary>Stable registration name for the CommonMark table markdown fallback.</summary>
    public const string CommonMarkTableMarkdownName = "CommonMark.Table.Markdown";
    /// <summary>Stable registration name for the CommonMark task-list markdown fallback.</summary>
    public const string CommonMarkTaskListMarkdownName = "CommonMark.TaskList.Markdown";
    /// <summary>Stable registration name for the CommonMark definition-list markdown fallback.</summary>
    public const string CommonMarkDefinitionListMarkdownName = "CommonMark.DefinitionList.Markdown";
    /// <summary>Stable registration name for the CommonMark footnote-definition markdown fallback.</summary>
    public const string CommonMarkFootnoteDefinitionMarkdownName = "CommonMark.FootnoteDefinition.Markdown";
    /// <summary>Stable registration name for the CommonMark details markdown fallback.</summary>
    public const string CommonMarkDetailsMarkdownName = "CommonMark.Details.Markdown";
    /// <summary>Stable registration name for the CommonMark custom-container markdown fallback.</summary>
    public const string CommonMarkCustomContainerMarkdownName = "CommonMark.CustomContainer.Markdown";
    /// <summary>Stable registration name for the CommonMark paragraph line-start markdown fallback.</summary>
    public const string CommonMarkParagraphLineStartMarkdownName = "CommonMark.ParagraphLineStart.Markdown";
    /// <summary>Stable registration name for the CommonMark attributed block markdown fallback.</summary>
    public const string CommonMarkAttributedBlockMarkdownName = "CommonMark.AttributedBlock.Markdown";
    /// <summary>Stable registration name for the GitHub definition-list markdown fallback.</summary>
    public const string GitHubDefinitionListMarkdownName = "GitHub.DefinitionList.Markdown";
    /// <summary>Stable registration name for the GitHub custom-container markdown fallback.</summary>
    public const string GitHubCustomContainerMarkdownName = "GitHub.CustomContainer.Markdown";
    /// <summary>Stable registration name for the GitHub paragraph line-start markdown fallback.</summary>
    public const string GitHubParagraphLineStartMarkdownName = "GitHub.ParagraphLineStart.Markdown";
    /// <summary>Stable registration name for the GitHub attributed block markdown fallback.</summary>
    public const string GitHubAttributedBlockMarkdownName = "GitHub.AttributedBlock.Markdown";

    /// <summary>Adds a portable markdown fallback for OfficeIMO callout blocks.</summary>
    public static void AddPortableCalloutMarkdownFallback(MarkdownWriteOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddIfMissing(options.BlockRenderExtensions, PortableCalloutMarkdownName, typeof(CalloutBlock), static (block, _) => {
            if (block is not CalloutBlock callout) {
                return null;
            }

            return RenderPortableCalloutMarkdown(callout);
        });
    }

    /// <summary>Adds a CommonMark-compatible markdown fallback for tables by emitting raw HTML tables.</summary>
    public static void AddCommonMarkTableMarkdownFallback(MarkdownWriteOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddIfMissing(options.BlockRenderExtensions, CommonMarkTableMarkdownName, typeof(TableBlock), static (block, _) => {
            if (block is not TableBlock table) {
                return null;
            }

            return ((IMarkdownBlock)table).RenderHtml();
        });
    }

    /// <summary>Adds a CommonMark-compatible markdown fallback for task lists by emitting raw HTML lists.</summary>
    public static void AddCommonMarkTaskListMarkdownFallback(MarkdownWriteOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddIfMissing(options.BlockRenderExtensions, CommonMarkTaskListMarkdownName, typeof(UnorderedListBlock), static (block, _) => {
            if (block is not UnorderedListBlock list || !ContainsTaskListItem(list.Items)) {
                return null;
            }

            return ((IMarkdownBlock)list).RenderHtml();
        });

        AddIfMissing(options.BlockRenderExtensions, CommonMarkTaskListMarkdownName + ".Ordered", typeof(OrderedListBlock), static (block, _) => {
            if (block is not OrderedListBlock list || !ContainsTaskListItem(list.Items)) {
                return null;
            }

            return ((IMarkdownBlock)list).RenderHtml();
        });
    }

    /// <summary>Adds a CommonMark-compatible markdown fallback for definition lists by emitting raw HTML.</summary>
    public static void AddCommonMarkDefinitionListMarkdownFallback(MarkdownWriteOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddDefinitionListHtmlMarkdownFallback(options, CommonMarkDefinitionListMarkdownName);
    }

    /// <summary>Adds a CommonMark-compatible markdown fallback for footnote definitions by emitting raw HTML.</summary>
    public static void AddCommonMarkFootnoteDefinitionMarkdownFallback(MarkdownWriteOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddIfMissing(options.BlockRenderExtensions, CommonMarkFootnoteDefinitionMarkdownName, typeof(FootnoteDefinitionBlock), static (block, _) => {
            if (block is not FootnoteDefinitionBlock footnote) {
                return null;
            }

            return ((IMarkdownBlock)footnote).RenderHtml();
        });
    }

    /// <summary>Adds a CommonMark-compatible markdown fallback for details blocks by emitting raw HTML.</summary>
    public static void AddCommonMarkDetailsMarkdownFallback(MarkdownWriteOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddIfMissing(options.BlockRenderExtensions, CommonMarkDetailsMarkdownName, typeof(DetailsBlock), static (block, _) => {
            if (block is not DetailsBlock details) {
                return null;
            }

            return ((IMarkdownBlock)details).RenderHtml();
        });
    }

    /// <summary>Adds a CommonMark-compatible markdown fallback for custom containers by emitting raw HTML.</summary>
    public static void AddCommonMarkCustomContainerMarkdownFallback(MarkdownWriteOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddCustomContainerHtmlMarkdownFallback(options, CommonMarkCustomContainerMarkdownName);
    }

    /// <summary>Adds CommonMark-compatible paragraph text escaping for line-start block markers.</summary>
    public static void AddCommonMarkParagraphLineStartMarkdownFallback(MarkdownWriteOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddParagraphLineStartMarkdownFallback(options, CommonMarkParagraphLineStartMarkdownName);
    }

    /// <summary>Adds a CommonMark-compatible markdown fallback for attributed blocks by emitting raw HTML.</summary>
    public static void AddCommonMarkAttributedBlockMarkdownFallback(MarkdownWriteOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddAttributedBlockHtmlMarkdownFallback(options, CommonMarkAttributedBlockMarkdownName);
    }

    /// <summary>Adds a GitHub-compatible markdown fallback for definition lists by emitting raw HTML.</summary>
    public static void AddGitHubDefinitionListMarkdownFallback(MarkdownWriteOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddDefinitionListHtmlMarkdownFallback(options, GitHubDefinitionListMarkdownName);
    }

    /// <summary>Adds a GitHub-compatible markdown fallback for custom containers by emitting raw HTML.</summary>
    public static void AddGitHubCustomContainerMarkdownFallback(MarkdownWriteOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddCustomContainerHtmlMarkdownFallback(options, GitHubCustomContainerMarkdownName);
    }

    /// <summary>Adds GitHub-compatible paragraph text escaping for line-start block markers.</summary>
    public static void AddGitHubParagraphLineStartMarkdownFallback(MarkdownWriteOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddParagraphLineStartMarkdownFallback(options, GitHubParagraphLineStartMarkdownName);
    }

    /// <summary>Adds a GitHub-compatible markdown fallback for attributed blocks by emitting raw HTML.</summary>
    public static void AddGitHubAttributedBlockMarkdownFallback(MarkdownWriteOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddAttributedBlockHtmlMarkdownFallback(options, GitHubAttributedBlockMarkdownName);
    }

    /// <summary>Adds a portable HTML fallback for OfficeIMO callout blocks.</summary>
    public static void AddPortableCalloutHtmlFallback(HtmlOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddIfMissing(options.BlockRenderExtensions, PortableCalloutHtmlName, typeof(CalloutBlock), static (block, options) => {
            if (block is not CalloutBlock callout) {
                return null;
            }

            return RenderPortableCalloutHtml(callout, options as HtmlOptions);
        });
    }

    /// <summary>Adds Markdig/GitHub-style HTML rendering for OfficeIMO callout blocks.</summary>
    public static void AddMarkdigAlertHtmlFallback(HtmlOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddIfMissing(options.BlockRenderExtensions, MarkdigAlertHtmlName, typeof(CalloutBlock), static (block, options) => {
            if (block is not CalloutBlock callout) {
                return null;
            }

            return RenderMarkdigAlertHtml(callout, options as HtmlOptions);
        });
    }

    /// <summary>
    /// Adds the full portable HTML fallback set: plain blockquotes for callouts, simple list-based TOC,
    /// and a simplified footnote section without OfficeIMO-specific chrome.
    /// </summary>
    public static void AddPortableHtmlFallbacks(HtmlOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddPortableCalloutHtmlFallback(options);
        AddPortableTocHtmlFallback(options);
        AddPortableFootnoteSectionHtmlFallback(options);
    }

    /// <summary>Adds a portable HTML fallback for TOC placeholders.</summary>
    public static void AddPortableTocHtmlFallback(HtmlOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        var prior = options.TocHtmlRenderer;
        options.TocHtmlRenderer = (tocOptions, entries, htmlOptions) =>
            prior?.Invoke(tocOptions, entries, htmlOptions) ?? RenderPortableTocHtml(tocOptions, entries, htmlOptions);
    }

    /// <summary>Adds a portable HTML fallback for the aggregated footnote section.</summary>
    public static void AddPortableFootnoteSectionHtmlFallback(HtmlOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        var prior = options.FootnoteSectionHtmlRenderer;
        options.FootnoteSectionHtmlRenderer = (footnotes, htmlOptions) =>
            prior?.Invoke(footnotes, htmlOptions) ?? RenderPortableFootnoteSectionHtml(footnotes);
    }

    private static void AddIfMissing<T>(
        List<T> extensions,
        string name,
        Type blockType,
        Func<IMarkdownBlock, object?, string?> renderer) where T : class {
        if (extensions == null) {
            throw new ArgumentNullException(nameof(extensions));
        }

        if (typeof(T) == typeof(MarkdownBlockMarkdownRenderExtension)) {
            if (extensions.Cast<MarkdownBlockMarkdownRenderExtension>().Any(extension =>
                string.Equals(extension.Name, name, StringComparison.OrdinalIgnoreCase))) {
                return;
            }

            extensions.Add((T)(object)new MarkdownBlockMarkdownRenderExtension(
                name,
                blockType,
                (block, options) => renderer(block, options)));
            return;
        }

        if (typeof(T) == typeof(MarkdownBlockHtmlRenderExtension)) {
            if (extensions.Cast<MarkdownBlockHtmlRenderExtension>().Any(extension =>
                string.Equals(extension.Name, name, StringComparison.OrdinalIgnoreCase))) {
                return;
            }

            extensions.Add((T)(object)new MarkdownBlockHtmlRenderExtension(
                name,
                blockType,
                (block, options) => renderer(block, options)));
        }
    }

    private static void AddDefinitionListHtmlMarkdownFallback(MarkdownWriteOptions options, string name) {
        AddIfMissing(options.BlockRenderExtensions, name, typeof(DefinitionListBlock), static (block, _) => {
            if (block is not DefinitionListBlock definitionList) {
                return null;
            }

            return ((IMarkdownBlock)definitionList).RenderHtml();
        });
    }

    private static void AddCustomContainerHtmlMarkdownFallback(MarkdownWriteOptions options, string name) {
        AddIfMissing(options.BlockRenderExtensions, name, typeof(CustomContainerBlock), static (block, _) => {
            if (block is not CustomContainerBlock container) {
                return null;
            }

            return ((IMarkdownBlock)container).RenderHtml();
        });
    }

    private static void AddParagraphLineStartMarkdownFallback(MarkdownWriteOptions options, string name) {
        AddIfMissing(options.BlockRenderExtensions, name, typeof(ParagraphBlock), static (block, _) => {
            if (block is not ParagraphBlock paragraph || !paragraph.Attributes.IsEmpty) {
                return null;
            }

            return paragraph.Inlines.RenderMarkdownWithTextEscaper(MarkdownEscaper.EscapeTextAndLineStarts);
        });
    }

    private static void AddAttributedBlockHtmlMarkdownFallback(MarkdownWriteOptions options, string name) {
        AddIfMissing(options.BlockRenderExtensions, name, typeof(MarkdownBlock), static (block, _) => {
            if (block is IMarkdownListBlock listBlock && ContainsAttributedListItem(listBlock.ListItems)) {
                return RenderListHtmlWithItemAttributes(listBlock);
            }

            if (block is MarkdownObject markdownObject && !markdownObject.Attributes.IsEmpty) {
                return block.RenderHtml();
            }

            return null;
        });
    }

    private static string RenderPortableCalloutMarkdown(CalloutBlock callout) {
        var titleMarkdown = callout.TitleInlines.RenderMarkdown();
        var visibleTitle = string.IsNullOrWhiteSpace(titleMarkdown)
            ? FormatTitleFromKind(callout.Kind)
            : titleMarkdown;

        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(visibleTitle)) {
            parts.Add("**" + visibleTitle.Trim() + "**");
        }

        if (callout.ChildBlocks.Count > 0) {
            for (int i = 0; i < callout.ChildBlocks.Count; i++) {
                var rendered = MarkdownBlockRenderDispatcher.RenderMarkdown(callout.ChildBlocks[i]);
                if (!string.IsNullOrWhiteSpace(rendered)) {
                    parts.Add(rendered!.TrimEnd());
                }
            }
        } else if (!string.IsNullOrWhiteSpace(callout.Body)) {
            parts.Add((callout.Body ?? string.Empty).TrimEnd());
        }

        return PrefixQuoteLines(string.Join("\n\n", parts));
    }

    private static bool ContainsTaskListItem(IReadOnlyList<ListItem> items) {
        if (items == null || items.Count == 0) {
            return false;
        }

        for (int i = 0; i < items.Count; i++) {
            var item = items[i];
            if (item == null) {
                continue;
            }

            if (item.IsTask) {
                return true;
            }

            for (int childIndex = 0; childIndex < item.Children.Count; childIndex++) {
                if (item.Children[childIndex] is IMarkdownListBlock childList
                    && ContainsTaskListItem(childList.ListItems)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool ContainsAttributedListItem(IReadOnlyList<ListItem> items) {
        if (items == null || items.Count == 0) {
            return false;
        }

        for (int i = 0; i < items.Count; i++) {
            var item = items[i];
            if (item == null) {
                continue;
            }

            if (!item.Attributes.IsEmpty) {
                return true;
            }

            for (int childIndex = 0; childIndex < item.Children.Count; childIndex++) {
                if (item.Children[childIndex] is IMarkdownListBlock childList
                    && ContainsAttributedListItem(childList.ListItems)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static string RenderListHtmlWithItemAttributes(IMarkdownListBlock listBlock) {
        using var _ = HtmlRenderContext.PushRenderListItemAttributes();
        return listBlock switch {
            UnorderedListBlock unordered => unordered.RenderHtml(renderItemAttributes: true),
            OrderedListBlock ordered => ordered.RenderHtml(renderItemAttributes: true),
            IMarkdownBlock block => block.RenderHtml(),
            _ => string.Empty
        };
    }

    private static string RenderPortableCalloutHtml(CalloutBlock callout, HtmlOptions? options) {
        var titleMarkdown = callout.TitleInlines.RenderMarkdown();
        var visibleTitle = string.IsNullOrWhiteSpace(titleMarkdown)
            ? FormatTitleFromKind(callout.Kind)
            : null;

        var sb = new StringBuilder();
        sb.Append("<blockquote>");

        if (!string.IsNullOrWhiteSpace(titleMarkdown)) {
            sb.Append("<p><strong>").Append(callout.TitleInlines.RenderHtml()).Append("</strong></p>");
        } else if (!string.IsNullOrWhiteSpace(visibleTitle)) {
            sb.Append("<p><strong>").Append(HtmlTextEncoder.Encode(visibleTitle, options)).Append("</strong></p>");
        }

        if (callout.ChildBlocks.Count > 0) {
            for (int i = 0; i < callout.ChildBlocks.Count; i++) {
                sb.Append(MarkdownBlockRenderDispatcher.RenderHtml(callout.ChildBlocks[i]));
            }
        } else {
            var lines = (callout.Body ?? string.Empty).Replace("\r\n", "\n").Split('\n');
            sb.Append("<p>");
            for (int i = 0; i < lines.Length; i++) {
                if (i > 0) {
                    sb.Append("<br/>");
                }
                sb.Append(HtmlTextEncoder.Encode(lines[i], options));
            }
            sb.Append("</p>");
        }

        sb.Append("</blockquote>");
        return sb.ToString();
    }

    private static string RenderMarkdigAlertHtml(CalloutBlock callout, HtmlOptions? options) {
        var kind = NormalizeAlertKind(callout.Kind);
        var defaultTitle = GetAlertTitle(kind);
        var explicitTitleMarkdown = callout.TitleInlines.RenderMarkdown();
        var hasExplicitTitle = !string.IsNullOrWhiteSpace(explicitTitleMarkdown);
        var sb = new StringBuilder();
        sb.Append("<div class=\"markdown-alert markdown-alert-")
            .Append(HtmlTextEncoder.Encode(kind, options))
            .Append("\">");

        if (hasExplicitTitle || !string.IsNullOrEmpty(defaultTitle)) {
            sb.Append("<p class=\"markdown-alert-title\">");
            if (!string.IsNullOrEmpty(defaultTitle)) {
                sb.Append(RenderAlertIcon(kind));
            }
            sb.Append(hasExplicitTitle
                    ? callout.TitleInlines.RenderHtml()
                    : HtmlTextEncoder.Encode(defaultTitle, options))
                .Append("</p>");
        }

        if (callout.ChildBlocks.Count > 0) {
            if (callout.ChildBlocks[0] is not ParagraphBlock) {
                sb.Append("<p></p>");
            }

            for (int i = 0; i < callout.ChildBlocks.Count; i++) {
                sb.Append(MarkdownBlockRenderDispatcher.RenderHtml(callout.ChildBlocks[i]));
            }
        } else if (!string.IsNullOrWhiteSpace(callout.Body)) {
            var lines = (callout.Body ?? string.Empty).Replace("\r\n", "\n").Split('\n');
            sb.Append("<p>");
            for (int i = 0; i < lines.Length; i++) {
                if (i > 0) {
                    sb.Append("<br/>");
                }
                sb.Append(HtmlTextEncoder.Encode(lines[i], options));
            }
            sb.Append("</p>");
        } else {
            sb.Append("<p></p>");
        }

        sb.Append("</div>");
        return sb.ToString();
    }

    private static string NormalizeAlertKind(string? kind) {
        var sourceKind = string.IsNullOrWhiteSpace(kind) ? "note" : kind!;
        var normalized = sourceKind.Trim().ToLowerInvariant();
        var builder = new StringBuilder(normalized.Length);
        for (int i = 0; i < normalized.Length; i++) {
            var ch = normalized[i];
            if ((ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9') || ch == '-' || ch == '_') {
                builder.Append(ch);
            }
        }

        return builder.Length == 0 ? "note" : builder.ToString();
    }

    private static string GetAlertTitle(string kind) =>
        kind switch {
            "note" => "Note",
            "tip" => "Tip",
            "important" => "Important",
            "warning" => "Warning",
            "caution" => "Caution",
            _ => string.Empty
        };

    private static string RenderAlertIcon(string kind) {
        var path = kind switch {
            "tip" => "M8.25 1.5a.75.75 0 0 0-1.5 0v1.25H5.5a.75.75 0 0 0 0 1.5h1.25V5.5a.75.75 0 0 0 1.5 0V4.25H9.5a.75.75 0 0 0 0-1.5H8.25V1.5ZM3.25 6.5a.75.75 0 0 0 0 1.5h9.5a.75.75 0 0 0 0-1.5h-9.5ZM4 10.25a.75.75 0 0 1 .75-.75h6.5a.75.75 0 0 1 0 1.5h-6.5a.75.75 0 0 1-.75-.75ZM5.75 12.5a.75.75 0 0 0 0 1.5h4.5a.75.75 0 0 0 0-1.5h-4.5Z",
            "important" => "M8 1.5a6.5 6.5 0 1 0 0 13 6.5 6.5 0 0 0 0-13ZM0 8a8 8 0 1 1 16 0A8 8 0 0 1 0 8Zm8-4a.75.75 0 0 1 .75.75v3.5a.75.75 0 0 1-1.5 0v-3.5A.75.75 0 0 1 8 4Zm0 8a1 1 0 1 1 0-2 1 1 0 0 1 0 2Z",
            "warning" => "M6.457 1.047c.659-1.234 2.427-1.234 3.086 0l6.082 11.378A1.75 1.75 0 0 1 14.082 15H1.918a1.75 1.75 0 0 1-1.543-2.575Zm1.763.707a.25.25 0 0 0-.44 0L1.698 13.132a.25.25 0 0 0 .22.368h12.164a.25.25 0 0 0 .22-.368Zm.53 3.996v2.5a.75.75 0 0 1-1.5 0v-2.5a.75.75 0 0 1 1.5 0ZM9 11a1 1 0 1 1-2 0 1 1 0 0 1 2 0Z",
            "caution" => "M6.457 1.047c.659-1.234 2.427-1.234 3.086 0l6.082 11.378A1.75 1.75 0 0 1 14.082 15H1.918a1.75 1.75 0 0 1-1.543-2.575Zm1.763.707a.25.25 0 0 0-.44 0L1.698 13.132a.25.25 0 0 0 .22.368h12.164a.25.25 0 0 0 .22-.368ZM8.75 4.75a.75.75 0 0 0-1.5 0v4.5a.75.75 0 0 0 1.5 0v-4.5ZM8 13a1 1 0 1 0 0-2 1 1 0 0 0 0 2Z",
            _ => "M0 8a8 8 0 1 1 16 0A8 8 0 0 1 0 8Zm8-6.5a6.5 6.5 0 1 0 0 13 6.5 6.5 0 0 0 0-13ZM6.5 7.75A.75.75 0 0 1 7.25 7h1a.75.75 0 0 1 .75.75v2.75h.25a.75.75 0 0 1 0 1.5h-2a.75.75 0 0 1 0-1.5h.25v-2h-.25a.75.75 0 0 1-.75-.75ZM8 6a1 1 0 1 1 0-2 1 1 0 0 1 0 2Z"
        };

        return "<svg viewBox=\"0 0 16 16\" version=\"1.1\" width=\"16\" height=\"16\" aria-hidden=\"true\"><path d=\"" + path + "\"></path></svg>";
    }

    private static string RenderPortableTocHtml(TocOptions options, IReadOnlyList<TocBlock.Entry> entries, HtmlOptions? htmlOptions) {
        if (entries == null || entries.Count == 0) {
            return string.Empty;
        }

        var toc = new TocBlock {
            Ordered = options.Ordered,
            NormalizeLevels = options.NormalizeToMinLevel
        };

        for (int i = 0; i < entries.Count; i++) {
            var entry = entries[i];
            toc.Entries.Add(new TocBlock.Entry {
                Level = entry.Level,
                Text = entry.Text,
                Anchor = entry.Anchor
            });
        }

        var bodyHtml = ((IMarkdownBlock)toc).RenderHtml();
        if (!options.IncludeTitle || string.IsNullOrWhiteSpace(options.Title)) {
            return bodyHtml;
        }

        int titleLevel = options.TitleLevel < 1 ? 1 : (options.TitleLevel > 6 ? 6 : options.TitleLevel);
        return $"<h{titleLevel}>{HtmlTextEncoder.Encode(options.Title, htmlOptions)}</h{titleLevel}>" + bodyHtml;
    }

    private static string RenderPortableFootnoteSectionHtml(IReadOnlyList<FootnoteDefinitionBlock> footnotes) {
        if (footnotes == null || footnotes.Count == 0) {
            return string.Empty;
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var sb = new StringBuilder();
        sb.Append("<section><hr />");

        for (int i = 0; i < footnotes.Count; i++) {
            var footnote = footnotes[i];
            if (footnote == null || string.IsNullOrWhiteSpace(footnote.Label) || !seen.Add(footnote.Label)) {
                continue;
            }

            sb.Append(((IMarkdownBlock)footnote).RenderHtml());
        }

        sb.Append("</section>");
        return sb.ToString();
    }

    private static string PrefixQuoteLines(string content) {
        if (string.IsNullOrWhiteSpace(content)) {
            return string.Empty;
        }

        var lines = content.Replace("\r\n", "\n").Split('\n');
        var sb = new StringBuilder();
        for (int i = 0; i < lines.Length; i++) {
            var line = lines[i];
            sb.Append('>');
            if (line.Length > 0) {
                sb.Append(' ').Append(line);
            }
            if (i < lines.Length - 1) {
                sb.AppendLine();
            }
        }

        return sb.ToString();
    }

    private static string FormatTitleFromKind(string kind) {
        if (string.IsNullOrWhiteSpace(kind)) {
            return string.Empty;
        }

        var value = kind.Trim();
        if (value.Length == 1) {
            return value.ToUpperInvariant();
        }

        return char.ToUpperInvariant(value[0]) + value.Substring(1).ToLowerInvariant();
    }
}
