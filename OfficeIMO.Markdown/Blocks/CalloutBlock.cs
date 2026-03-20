namespace OfficeIMO.Markdown;

/// <summary>
/// Docs/Markdown-style callout (admonition) block. Renders using
/// "> [!KIND] Title" followed by indented content lines.
/// </summary>
public sealed class CalloutBlock : MarkdownBlock, IMarkdownBlock, IChildMarkdownBlockContainer, ISyntaxChildrenMarkdownBlock, IOwnedSyntaxChildrenMarkdownBlock, ISyntaxMarkdownBlock {
    /// <summary>Admonition kind, e.g., info, warning, success.</summary>
    public string Kind { get; }
    private readonly string _fallbackBody;
    private readonly IReadOnlyList<IMarkdownBlock> _childBlocks;
    /// <summary>Callout title displayed inline with the marker.</summary>
    public string Title => InlinePlainText.Extract(TitleInlines);
    /// <summary>Parsed inline title content when available.</summary>
    public InlineSequence TitleInlines { get; }
    /// <summary>Callout body text (can include multiple lines). When parsed child blocks are available, this is derived from them.</summary>
    public string Body => _childBlocks.Count > 0 ? RenderBlocksAsBody(_childBlocks) : _fallbackBody;
    /// <summary>
    /// Parsed body blocks when the callout is created by the reader.
    /// This exposes callout content as owned child blocks for AST-style consumers.
    /// </summary>
    public IReadOnlyList<IMarkdownBlock> ChildBlocks => _childBlocks;

    /// <summary>
    /// Optional parsed body blocks. When present (produced by <see cref="MarkdownReader"/>),
    /// HTML/Markdown rendering uses these blocks instead of the raw <see cref="Body"/> string.
    /// </summary>
    internal IReadOnlyList<MarkdownSyntaxNode>? SyntaxChildren { get; }

    /// <summary>Creates a callout with the specified kind, title and body.</summary>
    public CalloutBlock(string kind, string title, string body)
        : this(kind, new InlineSequence().Text(title ?? string.Empty), body) {
    }

    internal CalloutBlock(string kind, InlineSequence titleInlines, string body) {
        Kind = (kind ?? "info").Trim();
        TitleInlines = titleInlines ?? new InlineSequence();
        _fallbackBody = body ?? string.Empty;
        _childBlocks = Array.Empty<IMarkdownBlock>();
    }

    internal CalloutBlock(string kind, string title, IReadOnlyList<IMarkdownBlock> children, IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren = null)
        : this(kind, new InlineSequence().Text(title ?? string.Empty), children, syntaxChildren) {
    }

    internal CalloutBlock(string kind, InlineSequence titleInlines, IReadOnlyList<IMarkdownBlock> children, IReadOnlyList<MarkdownSyntaxNode>? syntaxChildren = null) {
        Kind = (kind ?? "info").Trim();
        TitleInlines = titleInlines ?? new InlineSequence();
        _fallbackBody = string.Empty;
        _childBlocks = CopyChildren(children);
        SyntaxChildren = syntaxChildren;
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        string tag = Kind.ToUpperInvariant();
        StringBuilder sb = new StringBuilder();
        var titleMarkdown = TitleInlines.RenderMarkdown();
        if (string.IsNullOrWhiteSpace(titleMarkdown)) sb.AppendLine($"> [!{tag}]");
        else sb.AppendLine($"> [!{tag}] {titleMarkdown}");
        string bodyMarkdown;
        if (ChildBlocks.Count > 0) {
            var inner = new StringBuilder();
            for (int i = 0; i < ChildBlocks.Count; i++) {
                if (ChildBlocks[i] == null) continue;
                var rendered = ChildBlocks[i].RenderMarkdown();
                if (string.IsNullOrEmpty(rendered)) continue;
                inner.AppendLine(rendered.TrimEnd());
            }
            bodyMarkdown = inner.ToString().TrimEnd();
        } else {
            bodyMarkdown = Body ?? string.Empty;
        }
        foreach (string line in bodyMarkdown.Replace("\r\n", "\n").Split('\n')) {
            sb.AppendLine(line.Length == 0 ? ">" : ("> " + line));
        }
        return sb.ToString().TrimEnd();
    }

    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        var kind = System.Net.WebUtility.HtmlEncode(Kind);
        var titleMarkdown = TitleInlines.RenderMarkdown();
        var hasTitleInlines = !string.IsNullOrWhiteSpace(titleMarkdown);
        var titleText = hasTitleInlines ? TitleInlines.RenderHtml() : System.Net.WebUtility.HtmlEncode(FormatTitleFromKind(Kind));
        var hasVisibleTitle = hasTitleInlines || !string.IsNullOrWhiteSpace(FormatTitleFromKind(Kind));

        var sb = new StringBuilder();
        sb.Append("<blockquote class=\"callout ")
            .Append(kind)
            .Append("\" data-omd-callout-title-explicit=\"")
            .Append(hasTitleInlines ? "true" : "false")
            .Append("\">");
        if (hasVisibleTitle) {
            sb.Append("<p><strong>").Append(titleText).Append("</strong></p>");
        }

        if (ChildBlocks.Count > 0) {
            for (int i = 0; i < ChildBlocks.Count; i++) {
                if (ChildBlocks[i] == null) continue;
                sb.Append(ChildBlocks[i].RenderHtml());
            }
        } else {
            // Plain text body (builder-created callouts).
            var body = (Body ?? string.Empty).Replace("\r\n", "\n");
            var lines = body.Split('\n');
            sb.Append("<p>");
            for (int i = 0; i < lines.Length; i++) {
                if (i > 0) sb.Append("<br/>");
                sb.Append(System.Net.WebUtility.HtmlEncode(lines[i]));
            }
            sb.Append("</p>");
        }

        sb.Append("</blockquote>");
        return sb.ToString();
    }

    private static IReadOnlyList<IMarkdownBlock> CopyChildren(IReadOnlyList<IMarkdownBlock>? children) {
        if (children == null || children.Count == 0) {
            return Array.Empty<IMarkdownBlock>();
        }

        var copy = new IMarkdownBlock[children.Count];
        for (int i = 0; i < children.Count; i++) {
            copy[i] = children[i];
        }

        return copy;
    }

    private static string RenderBlocksAsBody(IReadOnlyList<IMarkdownBlock> blocks) {
        if (blocks == null || blocks.Count == 0) {
            return string.Empty;
        }

        var sb = new StringBuilder();
        for (int i = 0; i < blocks.Count; i++) {
            if (blocks[i] == null) {
                continue;
            }

            if (sb.Length > 0) {
                sb.Append("\n\n");
            }

            sb.Append((blocks[i].RenderMarkdown() ?? string.Empty)
                .Replace("\r\n", "\n")
                .Replace('\r', '\n')
                .TrimEnd());
        }

        return sb.ToString();
    }

    private static string FormatTitleFromKind(string kind) {
        if (string.IsNullOrWhiteSpace(kind)) return string.Empty;
        var t = kind.Trim();
        if (t.Length == 0) return string.Empty;
        if (t.Length == 1) return t.ToUpperInvariant();
        return char.ToUpperInvariant(t[0]) + t.Substring(1).ToLowerInvariant();
    }

    IReadOnlyList<MarkdownSyntaxNode>? ISyntaxChildrenMarkdownBlock.ProvidedSyntaxChildren => SyntaxChildren;

    IReadOnlyList<MarkdownSyntaxNode> IOwnedSyntaxChildrenMarkdownBlock.BuildOwnedSyntaxChildren() {
        if (SyntaxChildren != null && SyntaxChildren.Count > 0) {
            return SyntaxChildren;
        }

        return MarkdownBlockSyntaxBuilder.BuildChildSyntaxNodes(ChildBlocks);
    }

    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) {
        var calloutTitleMarkdown = TitleInlines.RenderMarkdown();
        return new MarkdownSyntaxNode(
            MarkdownSyntaxKind.Callout,
            span,
            string.IsNullOrWhiteSpace(calloutTitleMarkdown) ? Kind : Kind + ":" + calloutTitleMarkdown,
            ((IOwnedSyntaxChildrenMarkdownBlock)this).BuildOwnedSyntaxChildren(),
            this);
    }
}
