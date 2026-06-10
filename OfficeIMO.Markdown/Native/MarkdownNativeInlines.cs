namespace OfficeIMO.Markdown;

/// <summary>
/// AST-backed native projection for a markdown inline node.
/// </summary>
public sealed class MarkdownNativeInline {
    internal MarkdownNativeInline(
        MarkdownNativeInlineKind kind,
        MarkdownSyntaxNode syntaxNode,
        IReadOnlyList<MarkdownNativeInline> children,
        IReadOnlyList<MarkdownNativeInlineMetadata> metadata) {
        Kind = kind;
        SyntaxNode = syntaxNode ?? throw new ArgumentNullException(nameof(syntaxNode));
        SourceSpan = syntaxNode.SourceSpan;
        SourceInline = syntaxNode.AssociatedObject as IMarkdownInline;
        Literal = syntaxNode.Literal ?? string.Empty;
        Children = children ?? Array.Empty<MarkdownNativeInline>();
        Metadata = metadata ?? Array.Empty<MarkdownNativeInlineMetadata>();
        Text = ResolvePlainText(SourceInline, Literal, Children);
        Markdown = ResolveMarkdown(SourceInline, Literal);
        Id = MarkdownNativeInlineId.Create(kind, syntaxNode, SourceSpan);
    }

    /// <summary>Deterministic identity for this inline projection within stable markdown input.</summary>
    public string Id { get; }

    /// <summary>Native inline projection kind.</summary>
    public MarkdownNativeInlineKind Kind { get; }

    /// <summary>Syntax kind that produced this native inline.</summary>
    public MarkdownSyntaxKind SyntaxKind => SyntaxNode.Kind;

    /// <summary>Source span in the normalized markdown text when available.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }

    /// <summary>Syntax node that produced this native inline.</summary>
    public MarkdownSyntaxNode SyntaxNode { get; }

    /// <summary>Original OfficeIMO markdown inline backing this projection when available.</summary>
    public IMarkdownInline? SourceInline { get; }

    /// <summary>Literal syntax payload for leaf-like inline nodes.</summary>
    public string Literal { get; }

    /// <summary>Plain text represented by this inline node.</summary>
    public string Text { get; }

    /// <summary>Markdown represented by this inline node when renderable.</summary>
    public string Markdown { get; }

    /// <summary>Nested native inline nodes, excluding metadata leaves such as link target/title.</summary>
    public IReadOnlyList<MarkdownNativeInline> Children { get; }

    /// <summary>Source-backed metadata leaves such as link target/title or image alt/source.</summary>
    public IReadOnlyList<MarkdownNativeInlineMetadata> Metadata { get; }

    /// <summary>Returns the first metadata value with the supplied name.</summary>
    public string? GetMetadata(string name) {
        if (string.IsNullOrWhiteSpace(name) || Metadata.Count == 0) {
            return null;
        }

        for (var i = 0; i < Metadata.Count; i++) {
            if (string.Equals(Metadata[i].Name, name, StringComparison.OrdinalIgnoreCase)) {
                return Metadata[i].Value;
            }
        }

        return null;
    }

    /// <summary>Returns <see langword="true"/> when this inline's source span contains the supplied 1-based position.</summary>
    public bool ContainsPosition(int lineNumber, int columnNumber) =>
        SourceSpan.HasValue && SourceSpan.Value.ContainsPosition(lineNumber, columnNumber);

    private static string ResolvePlainText(IMarkdownInline? sourceInline, string literal, IReadOnlyList<MarkdownNativeInline> children) {
        if (sourceInline is IPlainTextMarkdownInline plainText) {
            var sb = new StringBuilder();
            plainText.AppendPlainText(sb);
            return sb.ToString();
        }

        if (children != null && children.Count > 0) {
            var sb = new StringBuilder();
            for (var i = 0; i < children.Count; i++) {
                sb.Append(children[i].Text);
            }

            return sb.ToString();
        }

        return literal ?? string.Empty;
    }

    private static string ResolveMarkdown(IMarkdownInline? sourceInline, string literal) {
        if (sourceInline is IRenderableMarkdownInline renderable) {
            return renderable.RenderMarkdown();
        }

        return literal ?? string.Empty;
    }
}

internal static class MarkdownNativeInlineProjection {
    internal static IReadOnlyList<MarkdownNativeInline> FromInlineContainer(MarkdownSyntaxNode? node) {
        if (node == null || node.Children.Count == 0) {
            return Array.Empty<MarkdownNativeInline>();
        }

        return FromSyntaxNodes(node.Children);
    }

    internal static IReadOnlyList<MarkdownNativeInline> FromInlineContainerChild(
        MarkdownSyntaxNode? node,
        MarkdownSyntaxKind childKind) {
        if (node == null || node.Children.Count == 0) {
            return Array.Empty<MarkdownNativeInline>();
        }

        for (var i = 0; i < node.Children.Count; i++) {
            if (node.Children[i].Kind == childKind) {
                return FromInlineContainer(node.Children[i]);
            }
        }

        return Array.Empty<MarkdownNativeInline>();
    }

    internal static IReadOnlyList<MarkdownNativeInline> FromFirstDescendantInlineContainer(MarkdownSyntaxNode? node) {
        var container = FindFirstInlineContainer(node);
        return container == null ? Array.Empty<MarkdownNativeInline>() : FromInlineContainer(container);
    }

    internal static IReadOnlyList<MarkdownNativeInline> FromListItemLeadContent(MarkdownSyntaxNode? node, ListItem? item) {
        if (node == null || item == null || item.Content.Nodes.Count == 0) {
            return Array.Empty<MarkdownNativeInline>();
        }

        for (var i = 0; i < node.Children.Count; i++) {
            if (node.Children[i].Kind == MarkdownSyntaxKind.Paragraph) {
                return FromInlineContainer(node.Children[i]);
            }
        }

        return Array.Empty<MarkdownNativeInline>();
    }

    private static IReadOnlyList<MarkdownNativeInline> FromSyntaxNodes(IReadOnlyList<MarkdownSyntaxNode> nodes) {
        if (nodes == null || nodes.Count == 0) {
            return Array.Empty<MarkdownNativeInline>();
        }

        var inlines = new List<MarkdownNativeInline>();
        for (var i = 0; i < nodes.Count; i++) {
            if (IsMetadataKind(nodes[i].Kind)) {
                continue;
            }

            if (!IsInlineKind(nodes[i].Kind)) {
                continue;
            }

            inlines.Add(Create(nodes[i]));
        }

        return inlines;
    }

    private static MarkdownNativeInline Create(MarkdownSyntaxNode node) {
        return new MarkdownNativeInline(
            MapKind(node.Kind),
            node,
            FromSyntaxNodes(node.Children),
            CreateMetadata(node.Children));
    }

    private static IReadOnlyList<MarkdownNativeInlineMetadata> CreateMetadata(IReadOnlyList<MarkdownSyntaxNode> children) {
        if (children == null || children.Count == 0) {
            return Array.Empty<MarkdownNativeInlineMetadata>();
        }

        var metadata = new List<MarkdownNativeInlineMetadata>();
        for (var i = 0; i < children.Count; i++) {
            if (TryGetMetadataName(children[i].Kind, out var name)) {
                metadata.Add(new MarkdownNativeInlineMetadata(name, children[i].Literal ?? string.Empty, children[i]));
            }
        }

        return metadata;
    }

    private static MarkdownSyntaxNode? FindFirstInlineContainer(MarkdownSyntaxNode? node) {
        if (node == null) {
            return null;
        }

        if (ContainsInlineChildren(node)) {
            return node;
        }

        for (var i = 0; i < node.Children.Count; i++) {
            var match = FindFirstInlineContainer(node.Children[i]);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    private static bool ContainsInlineChildren(MarkdownSyntaxNode node) {
        for (var i = 0; i < node.Children.Count; i++) {
            if (IsInlineKind(node.Children[i].Kind) && !IsMetadataKind(node.Children[i].Kind)) {
                return true;
            }
        }

        return false;
    }

    private static MarkdownNativeInlineKind MapKind(MarkdownSyntaxKind kind) {
        switch (kind) {
            case MarkdownSyntaxKind.InlineText:
                return MarkdownNativeInlineKind.Text;
            case MarkdownSyntaxKind.InlineCodeSpan:
                return MarkdownNativeInlineKind.Code;
            case MarkdownSyntaxKind.InlineLink:
                return MarkdownNativeInlineKind.Link;
            case MarkdownSyntaxKind.InlineImage:
                return MarkdownNativeInlineKind.Image;
            case MarkdownSyntaxKind.InlineImageLink:
                return MarkdownNativeInlineKind.ImageLink;
            case MarkdownSyntaxKind.InlineStrong:
                return MarkdownNativeInlineKind.Strong;
            case MarkdownSyntaxKind.InlineEmphasis:
                return MarkdownNativeInlineKind.Emphasis;
            case MarkdownSyntaxKind.InlineStrongEmphasis:
                return MarkdownNativeInlineKind.StrongEmphasis;
            case MarkdownSyntaxKind.InlineStrikethrough:
                return MarkdownNativeInlineKind.Strikethrough;
            case MarkdownSyntaxKind.InlineHighlight:
                return MarkdownNativeInlineKind.Highlight;
            case MarkdownSyntaxKind.InlineUnderline:
                return MarkdownNativeInlineKind.Underline;
            case MarkdownSyntaxKind.InlineHardBreak:
                return MarkdownNativeInlineKind.HardBreak;
            case MarkdownSyntaxKind.InlineHtmlTag:
                return MarkdownNativeInlineKind.HtmlTag;
            case MarkdownSyntaxKind.InlineHtmlRaw:
                return MarkdownNativeInlineKind.HtmlRaw;
            case MarkdownSyntaxKind.InlineFootnoteRef:
                return MarkdownNativeInlineKind.FootnoteRef;
            default:
                return MarkdownNativeInlineKind.Other;
        }
    }

    private static bool IsInlineKind(MarkdownSyntaxKind kind) {
        switch (kind) {
            case MarkdownSyntaxKind.InlineText:
            case MarkdownSyntaxKind.InlineCodeSpan:
            case MarkdownSyntaxKind.InlineLink:
            case MarkdownSyntaxKind.InlineImage:
            case MarkdownSyntaxKind.InlineImageLink:
            case MarkdownSyntaxKind.InlineStrong:
            case MarkdownSyntaxKind.InlineEmphasis:
            case MarkdownSyntaxKind.InlineStrongEmphasis:
            case MarkdownSyntaxKind.InlineStrikethrough:
            case MarkdownSyntaxKind.InlineHighlight:
            case MarkdownSyntaxKind.InlineUnderline:
            case MarkdownSyntaxKind.InlineHardBreak:
            case MarkdownSyntaxKind.InlineHtmlTag:
            case MarkdownSyntaxKind.InlineHtmlRaw:
            case MarkdownSyntaxKind.InlineFootnoteRef:
                return true;
            default:
                return false;
        }
    }

    private static bool IsMetadataKind(MarkdownSyntaxKind kind) => TryGetMetadataName(kind, out _);

    private static bool TryGetMetadataName(MarkdownSyntaxKind kind, out string name) {
        switch (kind) {
            case MarkdownSyntaxKind.InlineLinkTarget:
            case MarkdownSyntaxKind.ImageLinkTarget:
                name = "target";
                return true;
            case MarkdownSyntaxKind.InlineLinkTitle:
            case MarkdownSyntaxKind.ImageLinkTitle:
                name = "title";
                return true;
            case MarkdownSyntaxKind.InlineLinkHtmlTarget:
            case MarkdownSyntaxKind.ImageLinkHtmlTarget:
                name = "htmlTarget";
                return true;
            case MarkdownSyntaxKind.InlineLinkHtmlRel:
            case MarkdownSyntaxKind.ImageLinkHtmlRel:
                name = "htmlRel";
                return true;
            case MarkdownSyntaxKind.ImageAlt:
                name = "alt";
                return true;
            case MarkdownSyntaxKind.ImageSource:
                name = "source";
                return true;
            case MarkdownSyntaxKind.ImageTitle:
                name = "imageTitle";
                return true;
            default:
                name = string.Empty;
                return false;
        }
    }
}
