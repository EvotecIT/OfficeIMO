namespace OfficeIMO.Markdown;

/// <summary>
/// Built-in block render extension registrations for portable or host-specific output shaping.
/// </summary>
public static class MarkdownBlockRenderBuiltInExtensions {
    /// <summary>Stable registration name for the portable callout markdown fallback.</summary>
    public const string PortableCalloutMarkdownName = "Portable.Callout.Markdown";

    /// <summary>Stable registration name for the portable callout HTML fallback.</summary>
    public const string PortableCalloutHtmlName = "Portable.Callout.Html";
    /// <summary>Stable registration name for the portable TOC HTML fallback.</summary>
    public const string PortableTocHtmlName = "Portable.Toc.Html";
    /// <summary>Stable registration name for the portable footnote section HTML fallback.</summary>
    public const string PortableFootnoteSectionHtmlName = "Portable.Footnotes.Html";

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

    /// <summary>Adds a portable HTML fallback for OfficeIMO callout blocks.</summary>
    public static void AddPortableCalloutHtmlFallback(HtmlOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        AddIfMissing(options.BlockRenderExtensions, PortableCalloutHtmlName, typeof(CalloutBlock), static (block, _) => {
            if (block is not CalloutBlock callout) {
                return null;
            }

            return RenderPortableCalloutHtml(callout);
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
            prior?.Invoke(tocOptions, entries, htmlOptions) ?? RenderPortableTocHtml(tocOptions, entries);
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
                var rendered = callout.ChildBlocks[i]?.RenderMarkdown();
                if (!string.IsNullOrWhiteSpace(rendered)) {
                    parts.Add(rendered!.TrimEnd());
                }
            }
        } else if (!string.IsNullOrWhiteSpace(callout.Body)) {
            parts.Add((callout.Body ?? string.Empty).TrimEnd());
        }

        return PrefixQuoteLines(string.Join("\n\n", parts));
    }

    private static string RenderPortableCalloutHtml(CalloutBlock callout) {
        var titleMarkdown = callout.TitleInlines.RenderMarkdown();
        var visibleTitle = string.IsNullOrWhiteSpace(titleMarkdown)
            ? FormatTitleFromKind(callout.Kind)
            : null;

        var sb = new StringBuilder();
        sb.Append("<blockquote>");

        if (!string.IsNullOrWhiteSpace(titleMarkdown)) {
            sb.Append("<p><strong>").Append(callout.TitleInlines.RenderHtml()).Append("</strong></p>");
        } else if (!string.IsNullOrWhiteSpace(visibleTitle)) {
            sb.Append("<p><strong>").Append(System.Net.WebUtility.HtmlEncode(visibleTitle)).Append("</strong></p>");
        }

        if (callout.ChildBlocks.Count > 0) {
            for (int i = 0; i < callout.ChildBlocks.Count; i++) {
                sb.Append(callout.ChildBlocks[i]?.RenderHtml());
            }
        } else {
            var lines = (callout.Body ?? string.Empty).Replace("\r\n", "\n").Split('\n');
            sb.Append("<p>");
            for (int i = 0; i < lines.Length; i++) {
                if (i > 0) {
                    sb.Append("<br/>");
                }
                sb.Append(System.Net.WebUtility.HtmlEncode(lines[i]));
            }
            sb.Append("</p>");
        }

        sb.Append("</blockquote>");
        return sb.ToString();
    }

    private static string RenderPortableTocHtml(TocOptions options, IReadOnlyList<TocBlock.Entry> entries) {
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
        return $"<h{titleLevel}>{System.Net.WebUtility.HtmlEncode(options.Title)}</h{titleLevel}>" + bodyHtml;
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
