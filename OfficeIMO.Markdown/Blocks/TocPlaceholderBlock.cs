namespace OfficeIMO.Markdown;

/// <summary>
/// Placeholder that is replaced with a generated Table of Contents at render time.
/// </summary>
internal sealed class TocPlaceholderBlock : IMarkdownBlock, ISyntaxMarkdownBlock, ITocPlaceholderMarkdownBlock {
    public TocOptions Options { get; }
    public TocPlaceholderBlock(TocOptions options) { Options = options; }
    string IMarkdownBlock.RenderMarkdown() => string.Empty; // Replaced during render
    string IMarkdownBlock.RenderHtml() => string.Empty; // Replaced during render

    bool IBodySidebarMarkdownBlock.UsesSidebarLayout() =>
        Options.Layout == TocLayout.SidebarLeft || Options.Layout == TocLayout.SidebarRight;

    bool IBodySidebarMarkdownBlock.SuppressesPrecedingHeadingTitle() =>
        Options.IncludeTitle &&
        (Options.Layout == TocLayout.SidebarLeft ||
         Options.Layout == TocLayout.SidebarRight ||
         Options.Layout == TocLayout.Panel);

    bool ITocPlaceholderMarkdownBlock.RequiresScrollSpy() => Options.ScrollSpy;

    string IBodySidebarMarkdownBlock.WrapSidebarLayoutHtml(string navigationHtml, string contentHtml) {
        var side = Options.Layout == TocLayout.SidebarLeft ? "left" : "right";
        string widthStyle = Options.WidthPx.HasValue ? $" style=\"--md-toc-width: {Options.WidthPx.Value}px\"" : string.Empty;
        var sb = new System.Text.StringBuilder();
        sb.Append($"<div class=\"md-layout two-col {side}\"{widthStyle}>");
        if (side == "left") {
            sb.Append(navigationHtml).Append("<div class=\"md-content\">").Append(contentHtml).Append("</div>");
        } else {
            sb.Append("<div class=\"md-content\">").Append(contentHtml).Append("</div>").Append(navigationHtml);
        }
        sb.Append("</div>");
        return sb.ToString();
    }

    TocBlock ITocPlaceholderMarkdownBlock.RealizeToc(IReadOnlyList<IMarkdownBlock> blocks, int placeholderIndex, MarkdownHeadingCatalog headingCatalog) {
        var toc = new TocBlock { Ordered = Options.Ordered, NormalizeLevels = Options.NormalizeToMinLevel };
        string? titleAnchor = headingCatalog.GetPrecedingHeadingAnchor(blocks, placeholderIndex, Options);
        foreach (var entry in headingCatalog.BuildTocEntries(blocks, placeholderIndex, Options, titleAnchor)) {
            toc.Entries.Add(entry);
        }
        return toc;
    }

    string IContextualHtmlMarkdownBlock.RenderHtml(MarkdownBodyRenderContext context) {
        int titleLevel = Options.TitleLevel < 1 ? 1 : (Options.TitleLevel > 6 ? 6 : Options.TitleLevel);
        int placeholderIndex = System.Array.IndexOf(context.Blocks.ToArray(), this);
        string? titleAnchor = context.HeadingCatalog.GetPrecedingHeadingAnchor(context.Blocks, placeholderIndex, Options);
        var relevant = context.HeadingCatalog.BuildTocEntries(context.Blocks, placeholderIndex, Options, titleAnchor)
            .Select(e => (e.Level, e.Text, e.Anchor))
            .ToList();
        if (relevant.Count == 0) return string.Empty;

        var listTag = Options.Ordered ? "ol" : "ul";
        var sbNested = new System.Text.StringBuilder(relevant.Count * 64);
        int baseLevel = (Options.NormalizeToMinLevel ? relevant.Min(r => r.Level) : 1);
        int currentDepth = 0;
        bool any = false;
        sbNested.Append('<').Append(listTag).Append('>');
        for (int i = 0; i < relevant.Count; i++) {
            var e = relevant[i];
            int depth = System.Math.Max(0, e.Level - baseLevel);

            if (any) {
                if (depth == currentDepth) {
                    sbNested.Append("</li>");
                } else if (depth > currentDepth) {
                    for (int d = currentDepth; d < depth; d++) sbNested.Append('<').Append(listTag).Append('>');
                } else {
                    for (int d = currentDepth; d > depth; d--) sbNested.Append("</li></").Append(listTag).Append('>');
                    sbNested.Append("</li>");
                }
            }

            sbNested.Append("<li><a href=\"")
                .Append('#').Append(System.Net.WebUtility.HtmlEncode(e.Anchor))
                .Append("\">")
                .Append(System.Net.WebUtility.HtmlEncode(e.Text))
                .Append("</a>");

            currentDepth = depth;
            any = true;
        }
        if (any) sbNested.Append("</li>");
        for (int d = currentDepth; d > 0; d--) sbNested.Append("</").Append(listTag).Append("></li>");
        sbNested.Append("</").Append(listTag).Append('>');

        if (Options.Collapsible) {
            string open = Options.Collapsed ? string.Empty : " open";
            string summary = System.Net.WebUtility.HtmlEncode(Options.Title ?? "Contents");
            var sbWrap = new System.Text.StringBuilder();
            sbWrap.Append("<details class=\"md-toc\"").Append(open).Append("><summary>")
                .Append(summary).Append("</summary>")
                .Append(sbNested.ToString())
                .Append("</details>");
            return sbWrap.ToString();
        }

        bool enhanced = Options.Layout != TocLayout.List || Options.ScrollSpy || Options.Sticky;
        if (enhanced) {
            var classes = new System.Text.StringBuilder("md-toc");
            if (Options.Layout == TocLayout.Panel) classes.Append(" panel");
            if (Options.Layout == TocLayout.SidebarRight) classes.Append(" sidebar right");
            if (Options.Layout == TocLayout.SidebarLeft) classes.Append(" sidebar left");
            if (Options.Sticky) classes.Append(" sticky");
            if (Options.ScrollSpy) classes.Append(" md-scrollspy autoscroll");
            if (Options.Chrome == TocChrome.None) classes.Append(" no-chrome");
            if (Options.Chrome == TocChrome.Outline) classes.Append(" outline");
            if (Options.Chrome == TocChrome.Panel) classes.Append(" panel");
            if (Options.HideOnNarrow) classes.Append(" hide-narrow");
            string aria = System.Net.WebUtility.HtmlEncode(Options.AriaLabel ?? "Table of Contents");
            var sb = new System.Text.StringBuilder();
            sb.Append("<nav role=\"navigation\" aria-label=\"").Append(aria).Append("\" class=\"").Append(classes).Append("\"");
            if (Options.ScrollSpy) sb.Append(" data-md-scrollspy=\"1\"");
            if (Options.Sticky) sb.Append(" data-autoscroll=\"1\"");
            sb.Append(">");
            if (Options.IncludeTitle && !string.IsNullOrWhiteSpace(Options.Title)) {
                sb.Append("<div class=\"toc-title\">").Append(System.Net.WebUtility.HtmlEncode(Options.Title)).Append("</div>");
            }
            sb.Append(sbNested.ToString());
            sb.Append("</nav>");
            return sb.ToString();
        }

        if (Options.IncludeTitle) {
            var sbo = new System.Text.StringBuilder();
            sbo.Append("<h").Append(titleLevel).Append('>')
               .Append(System.Net.WebUtility.HtmlEncode(Options.Title))
               .Append("</h").Append(titleLevel).Append('>')
               .Append(sbNested.ToString());
            return sbo.ToString();
        }
        return sbNested.ToString();
    }
    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(MarkdownSyntaxKind.TocPlaceholder, span);
}
