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
        SourceSpan = ResolveNativeSourceSpan(syntaxNode);
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

    private static MarkdownSourceSpan? ResolveNativeSourceSpan(MarkdownSyntaxNode syntaxNode) {
        if (!UsesContentSpanForNativeInline(syntaxNode.Kind) || syntaxNode.Children.Count == 0) {
            return syntaxNode.SourceSpan;
        }

        var contentChildren = syntaxNode.Children
            .Where(child => !IsInlineMarkerKind(child.Kind))
            .ToArray();
        return MarkdownBlockSyntaxBuilder.GetAggregateSpan(contentChildren) ?? syntaxNode.SourceSpan;
    }

    private static bool UsesContentSpanForNativeInline(MarkdownSyntaxKind kind) =>
        kind == MarkdownSyntaxKind.InlineStrong
        || kind == MarkdownSyntaxKind.InlineEmphasis
        || kind == MarkdownSyntaxKind.InlineStrongEmphasis
        || kind == MarkdownSyntaxKind.InlineStrikethrough
        || kind == MarkdownSyntaxKind.InlineHighlight
        || kind == MarkdownSyntaxKind.InlineInserted
        || kind == MarkdownSyntaxKind.InlineSuperscript
        || kind == MarkdownSyntaxKind.InlineSubscript;

    private static bool IsInlineMarkerKind(MarkdownSyntaxKind kind) =>
        kind == MarkdownSyntaxKind.InlineOpeningMarker
        || kind == MarkdownSyntaxKind.InlineSeparatorMarker
        || kind == MarkdownSyntaxKind.InlineClosingMarker;
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

    internal static IReadOnlyList<MarkdownNativeInline> FromTableCellDirectContent(MarkdownSyntaxNode? node) {
        if (node == null) {
            return Array.Empty<MarkdownNativeInline>();
        }

        if (ContainsInlineChildren(node)) {
            return FromInlineContainer(node);
        }

        for (var i = 0; i < node.Children.Count; i++) {
            if (node.Children[i].Kind == MarkdownSyntaxKind.Paragraph) {
                return FromInlineContainer(node.Children[i]);
            }
        }

        return Array.Empty<MarkdownNativeInline>();
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

            if (!IsInlineNode(nodes[i])) {
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
            CreateMetadata(node));
    }

    private static IReadOnlyList<MarkdownNativeInlineMetadata> CreateMetadata(MarkdownSyntaxNode node) {
        if (node == null) {
            return Array.Empty<MarkdownNativeInlineMetadata>();
        }

        var children = node.Children;
        var metadata = new List<MarkdownNativeInlineMetadata>();
        if (children != null) {
            for (var i = 0; i < children.Count; i++) {
                if (TryGetMetadataName(children[i].Kind, out var name)) {
                    metadata.Add(new MarkdownNativeInlineMetadata(name, children[i].Literal ?? string.Empty, children[i]));
                }
            }
        }

        AddFormattingMarkerMetadata(node, metadata);
        AddCodeSpanContentMetadata(node, metadata);
        AddEscapedTextMetadata(node, metadata);
        AddDecodedEntityMetadata(node, metadata);
        AddHardBreakMarkerMetadata(node, metadata);
        AddAbbreviationMetadata(node, metadata);
        AddGenericAttributeMetadata(node, metadata);

        if (metadata.Count == 0) {
            return Array.Empty<MarkdownNativeInlineMetadata>();
        }

        return metadata;
    }

    private static void AddFormattingMarkerMetadata(MarkdownSyntaxNode node, List<MarkdownNativeInlineMetadata> metadata) {
        if (node.AssociatedObject is not MarkdownInline inline) {
            return;
        }

        var openingMarkerSpan = MarkdownInlineMetadataSourceSpans.GetOpeningMarkerSpan(inline);
        if (openingMarkerSpan.HasValue) {
            metadata.Add(new MarkdownNativeInlineMetadata(
                "openingMarker",
                MarkdownInlineMetadataSourceSpans.GetOpeningMarker(inline) ?? string.Empty,
                node,
                openingMarkerSpan));
        }

        var separatorMarkerSpan = MarkdownInlineMetadataSourceSpans.GetSeparatorMarkerSpan(inline);
        if (separatorMarkerSpan.HasValue) {
            metadata.Add(new MarkdownNativeInlineMetadata(
                "separatorMarker",
                MarkdownInlineMetadataSourceSpans.GetSeparatorMarker(inline) ?? string.Empty,
                node,
                separatorMarkerSpan));
        }

        var closingMarkerSpan = MarkdownInlineMetadataSourceSpans.GetClosingMarkerSpan(inline);
        if (closingMarkerSpan.HasValue) {
            metadata.Add(new MarkdownNativeInlineMetadata(
                "closingMarker",
                MarkdownInlineMetadataSourceSpans.GetClosingMarker(inline) ?? string.Empty,
                node,
                closingMarkerSpan));
        }
    }

    private static void AddCodeSpanContentMetadata(MarkdownSyntaxNode node, List<MarkdownNativeInlineMetadata> metadata) {
        if (node.AssociatedObject is not CodeSpanInline codeSpan) {
            return;
        }

        var contentSpan = MarkdownInlineMetadataSourceSpans.GetCodeSpanContentSpan(codeSpan);
        if (contentSpan.HasValue) {
            metadata.Add(new MarkdownNativeInlineMetadata(
                "content",
                codeSpan.Text,
                node,
                contentSpan));
        }
    }

    private static void AddEscapedTextMetadata(MarkdownSyntaxNode node, List<MarkdownNativeInlineMetadata> metadata) {
        if (node.AssociatedObject is not TextRun text) {
            return;
        }

        var escapeMarkerSpan = MarkdownInlineMetadataSourceSpans.GetEscapeMarkerSpan(text);
        if (escapeMarkerSpan.HasValue) {
            metadata.Add(new MarkdownNativeInlineMetadata(
                "escapeMarker",
                MarkdownInlineMetadataSourceSpans.GetEscapeMarker(text) ?? string.Empty,
                node,
                escapeMarkerSpan));
        }

        var escapedCharacterSpan = MarkdownInlineMetadataSourceSpans.GetEscapedCharacterSpan(text);
        if (escapedCharacterSpan.HasValue) {
            metadata.Add(new MarkdownNativeInlineMetadata(
                "escapedCharacter",
                MarkdownInlineMetadataSourceSpans.GetEscapedCharacter(text) ?? string.Empty,
                node,
                escapedCharacterSpan));
        }
    }

    private static void AddDecodedEntityMetadata(MarkdownSyntaxNode node, List<MarkdownNativeInlineMetadata> metadata) {
        if (node.AssociatedObject is not DecodedHtmlEntityTextRun decodedEntity) {
            return;
        }

        var sourceTextSpan = MarkdownInlineMetadataSourceSpans.GetDecodedEntitySourceTextSpan(decodedEntity);
        if (sourceTextSpan.HasValue) {
            metadata.Add(new MarkdownNativeInlineMetadata(
                "sourceText",
                MarkdownInlineMetadataSourceSpans.GetDecodedEntitySourceText(decodedEntity) ?? string.Empty,
                node,
                sourceTextSpan));
        }
    }

    private static void AddHardBreakMarkerMetadata(MarkdownSyntaxNode node, List<MarkdownNativeInlineMetadata> metadata) {
        if (node.AssociatedObject is not HardBreakInline hardBreak) {
            return;
        }

        var markerSpan = MarkdownInlineMetadataSourceSpans.GetHardBreakMarkerSpan(hardBreak);
        if (markerSpan.HasValue) {
            metadata.Add(new MarkdownNativeInlineMetadata(
                "marker",
                MarkdownInlineMetadataSourceSpans.GetHardBreakMarker(hardBreak) ?? string.Empty,
                node,
                markerSpan));
        }
    }

    private static void AddAbbreviationMetadata(MarkdownSyntaxNode node, List<MarkdownNativeInlineMetadata> metadata) {
        if (node.AssociatedObject is not AbbreviationInline abbreviation) {
            return;
        }

        RemoveMetadata(metadata, "text");
        RemoveMetadata(metadata, "title");

        metadata.Add(new MarkdownNativeInlineMetadata(
            "text",
            abbreviation.Text,
            node,
            MarkdownInlineMetadataSourceSpans.GetAbbreviationTextSpan(abbreviation) ?? node.SourceSpan));

        metadata.Add(new MarkdownNativeInlineMetadata(
            "title",
            abbreviation.Title,
            node,
            MarkdownInlineMetadataSourceSpans.GetAbbreviationTitleSpan(abbreviation)));
    }

    private static void AddGenericAttributeMetadata(MarkdownSyntaxNode node, List<MarkdownNativeInlineMetadata> metadata) {
        if (node.AssociatedObject is not MarkdownObject markdownObject || markdownObject.Attributes.IsEmpty) {
            return;
        }

        var sourceSpan = MarkdownGenericAttributeSourceSpans.GetSourceSpan(markdownObject);
        if (!sourceSpan.HasValue) {
            return;
        }

        metadata.Add(new MarkdownNativeInlineMetadata(
            "attributes",
            MarkdownGenericAttributeSourceSpans.GetSourceText(markdownObject) ?? string.Empty,
            node,
            sourceSpan));
    }

    private static void RemoveMetadata(List<MarkdownNativeInlineMetadata> metadata, string name) {
        for (var i = metadata.Count - 1; i >= 0; i--) {
            if (string.Equals(metadata[i].Name, name, StringComparison.OrdinalIgnoreCase)) {
                metadata.RemoveAt(i);
            }
        }
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
            if (IsInlineNode(node.Children[i]) && !IsMetadataKind(node.Children[i].Kind)) {
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
            case MarkdownSyntaxKind.InlineInserted:
                return MarkdownNativeInlineKind.Inserted;
            case MarkdownSyntaxKind.InlineSuperscript:
                return MarkdownNativeInlineKind.Superscript;
            case MarkdownSyntaxKind.InlineSubscript:
                return MarkdownNativeInlineKind.Subscript;
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
            case MarkdownSyntaxKind.InlineAbbreviation:
                return MarkdownNativeInlineKind.Abbreviation;
            default:
                return MarkdownNativeInlineKind.Other;
        }
    }

    private static bool IsInlineNode(MarkdownSyntaxNode node) =>
        node != null && (IsInlineKind(node.Kind) || (node.Kind == MarkdownSyntaxKind.Unknown && node.AssociatedObject is IMarkdownInline));

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
            case MarkdownSyntaxKind.InlineInserted:
            case MarkdownSyntaxKind.InlineSuperscript:
            case MarkdownSyntaxKind.InlineSubscript:
            case MarkdownSyntaxKind.InlineUnderline:
            case MarkdownSyntaxKind.InlineHardBreak:
            case MarkdownSyntaxKind.InlineHtmlTag:
            case MarkdownSyntaxKind.InlineHtmlRaw:
            case MarkdownSyntaxKind.InlineFootnoteRef:
            case MarkdownSyntaxKind.InlineAbbreviation:
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
            case MarkdownSyntaxKind.InlineFootnoteLabel:
                name = "label";
                return true;
            case MarkdownSyntaxKind.InlineAbbreviationTitle:
                name = "title";
                return true;
            case MarkdownSyntaxKind.InlineAbbreviationText:
                name = "text";
                return true;
            default:
                name = string.Empty;
                return false;
        }
    }
}
