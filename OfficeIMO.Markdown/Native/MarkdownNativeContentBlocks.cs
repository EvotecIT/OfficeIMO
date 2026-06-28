namespace OfficeIMO.Markdown;

/// <summary>
/// Native projection for a heading block.
/// </summary>
public sealed class MarkdownNativeHeadingBlock : MarkdownNativeBlock {
    internal MarkdownNativeHeadingBlock(HeadingBlock heading, MarkdownSyntaxNode syntaxNode)
        : base(MarkdownNativeBlockKind.Heading, heading, syntaxNode) {
        Heading = heading;
        Level = heading.Level;
        Inlines = heading.Inlines;
        InlineRuns = MarkdownNativeInlineProjection.FromInlineContainerChild(syntaxNode, MarkdownSyntaxKind.HeadingText);
        Text = heading.Text;
        LevelSourceSpan = GetChildSpan(syntaxNode, MarkdownSyntaxKind.HeadingLevel);
        TextSourceSpan = GetChildSpan(syntaxNode, MarkdownSyntaxKind.HeadingText);
        ClosingMarkerSourceSpan = GetChildSpan(syntaxNode, MarkdownSyntaxKind.HeadingClosingMarker);
        ClosingMarkerText = heading.ClosingMarkerText;
    }

    /// <summary>Source heading block.</summary>
    public HeadingBlock Heading { get; }

    /// <summary>Heading level, where 1 is H1.</summary>
    public int Level { get; }

    /// <summary>Plain-text heading content.</summary>
    public string Text { get; }

    /// <summary>Structured inline nodes.</summary>
    public InlineSequence Inlines { get; }

    /// <summary>AST-backed native inline projection with source spans.</summary>
    public IReadOnlyList<MarkdownNativeInline> InlineRuns { get; }

    /// <summary>Source span for the heading marker or setext underline that determines the level.</summary>
    public MarkdownSourceSpan? LevelSourceSpan { get; }

    /// <summary>Source span for the heading text payload.</summary>
    public MarkdownSourceSpan? TextSourceSpan { get; }

    /// <summary>Source span for the optional ATX closing marker token.</summary>
    public MarkdownSourceSpan? ClosingMarkerSourceSpan { get; }

    /// <summary>Exact optional ATX closing marker token when parsed from markdown.</summary>
    public string? ClosingMarkerText { get; }

    private static MarkdownSourceSpan? GetChildSpan(MarkdownSyntaxNode syntaxNode, MarkdownSyntaxKind kind) =>
        syntaxNode?.Children.FirstOrDefault(child => child.Kind == kind)?.SourceSpan;
}

/// <summary>
/// Native projection for an ordered or unordered list block.
/// </summary>
public sealed class MarkdownNativeListBlock : MarkdownNativeBlock {
    internal MarkdownNativeListBlock(
        IMarkdownListBlock list,
        MarkdownSyntaxNode syntaxNode,
        IReadOnlyList<MarkdownNativeListItem> items)
        : base(MarkdownNativeBlockKind.List, list, syntaxNode) {
        List = list;
        Items = items ?? Array.Empty<MarkdownNativeListItem>();
        IsOrdered = list is OrderedListBlock;
        Start = list is OrderedListBlock ordered ? ordered.Start : null;
    }

    /// <summary>Source list block.</summary>
    public IMarkdownBlock List { get; }

    /// <summary>Whether the list is ordered.</summary>
    public bool IsOrdered { get; }

    /// <summary>Ordered-list start value, or <c>null</c> for unordered lists.</summary>
    public int? Start { get; }

    /// <summary>Native list items in document order.</summary>
    public IReadOnlyList<MarkdownNativeListItem> Items { get; }
}

/// <summary>
/// Native projection for a list item.
/// </summary>
public sealed class MarkdownNativeListItem {
    internal MarkdownNativeListItem(
        ListItem item,
        MarkdownSyntaxNode syntaxNode,
        IReadOnlyList<MarkdownNativeBlock> children) {
        Item = item ?? throw new ArgumentNullException(nameof(item));
        SyntaxNode = syntaxNode ?? throw new ArgumentNullException(nameof(syntaxNode));
        SourceSpan = syntaxNode.SourceSpan ?? item.SourceSpan;
        ContentSourceSpan = GetContentSourceSpan(syntaxNode);
        Children = children ?? Array.Empty<MarkdownNativeBlock>();
        Text = InlinePlainText.Extract(item.Content);
        Inlines = item.Content;
        InlineRuns = MarkdownNativeInlineProjection.FromListItemLeadContent(syntaxNode, item);
        AdditionalParagraphs = item.AdditionalParagraphs;
        IsTask = item.IsTask;
        Checked = item.Checked;
        MarkerSourceSpan = item.MarkerSourceSpan;
        MarkerText = item.MarkerText;
        TaskMarkerSourceSpan = item.TaskMarkerSourceSpan;
        TaskMarkerText = item.TaskMarkerText;
        Level = item.Level;
        Id = MarkdownNativeListItemId.Create(item, syntaxNode, SourceSpan);
    }

    /// <summary>Deterministic identity for this list item within stable markdown input.</summary>
    public string Id { get; }

    /// <summary>Source list item.</summary>
    public ListItem Item { get; }

    /// <summary>Syntax node that produced this list item.</summary>
    public MarkdownSyntaxNode SyntaxNode { get; }

    /// <summary>Full list-item source span in the normalized markdown text when available.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }

    /// <summary>Source span for the list-item content, excluding list and task marker tokens, when available.</summary>
    public MarkdownSourceSpan? ContentSourceSpan { get; }

    /// <summary>Plain-text lead content.</summary>
    public string Text { get; }

    /// <summary>Structured lead inline nodes.</summary>
    public InlineSequence Inlines { get; }

    /// <summary>AST-backed native inline projection for the lead content.</summary>
    public IReadOnlyList<MarkdownNativeInline> InlineRuns { get; }

    /// <summary>Additional paragraph inline nodes owned by this list item.</summary>
    public IReadOnlyList<InlineSequence> AdditionalParagraphs { get; }

    /// <summary>Nested native blocks, including lead paragraph blocks when present in the syntax tree.</summary>
    public IReadOnlyList<MarkdownNativeBlock> Children { get; }

    /// <summary>Whether this list item is a task item.</summary>
    public bool IsTask { get; }

    /// <summary>Whether this task item is checked.</summary>
    public bool Checked { get; }

    /// <summary>Source span for the list marker token when this item was parsed from markdown.</summary>
    public MarkdownSourceSpan? MarkerSourceSpan { get; }

    /// <summary>Exact list marker token when this item was parsed from markdown.</summary>
    public string? MarkerText { get; }

    /// <summary>Source span for the task marker token when this item was parsed from markdown.</summary>
    public MarkdownSourceSpan? TaskMarkerSourceSpan { get; }

    /// <summary>Exact task marker token when this item was parsed from markdown.</summary>
    public string? TaskMarkerText { get; }

    /// <summary>Indentation level from the source list item.</summary>
    public int Level { get; }

    private static MarkdownSourceSpan? GetContentSourceSpan(MarkdownSyntaxNode syntaxNode) {
        var children = new List<MarkdownSyntaxNode>();
        for (var i = 0; i < syntaxNode.Children.Count; i++) {
            var child = syntaxNode.Children[i];
            if (child.Kind == MarkdownSyntaxKind.ListMarker || child.Kind == MarkdownSyntaxKind.TaskListMarker) {
                continue;
            }

            children.Add(child);
        }

        return MarkdownBlockSyntaxBuilder.GetAggregateSpan(children);
    }
}

internal static class MarkdownNativeListItemId {
    internal static string Create(ListItem item, MarkdownSyntaxNode syntaxNode, MarkdownSourceSpan? sourceSpan) {
        var span = sourceSpan.HasValue ? sourceSpan.Value.ToString() : "nosource";
        var literal = syntaxNode.Literal ?? item.RenderMarkdown() ?? string.Empty;
        var path = MarkdownNativeBlockId.BuildSyntaxPath(syntaxNode);
        var key = "ListItem|" + syntaxNode.Kind + "|" + span + "|" + path + "|" + literal;
        return "mdn-li-" + ComputeFnv1A64(key).ToString("x16", System.Globalization.CultureInfo.InvariantCulture);
    }

    private static ulong ComputeFnv1A64(string value) {
        const ulong offsetBasis = 14695981039346656037UL;
        const ulong prime = 1099511628211UL;

        var hash = offsetBasis;
        for (var i = 0; i < value.Length; i++) {
            hash ^= value[i];
            hash *= prime;
        }

        return hash;
    }
}

/// <summary>
/// Native projection for an image block.
/// </summary>
public sealed class MarkdownNativeImageBlock : MarkdownNativeBlock {
    internal MarkdownNativeImageBlock(ImageBlock image, MarkdownSyntaxNode syntaxNode)
        : base(MarkdownNativeBlockKind.Image, image, syntaxNode) {
        Image = image;
        Source = image.Path;
        Alt = image.Alt;
        PlainAlt = image.PlainAlt;
        Title = image.Title;
        Width = image.Width;
        Height = image.Height;
        Caption = image.Caption;
        LinkUrl = image.LinkUrl;
        LinkTitle = image.LinkTitle;
        LinkTarget = image.LinkTarget;
        LinkRel = image.LinkRel;
        PictureFallbackPath = image.PictureFallbackPath;
        PictureSources = image.PictureSources.ToArray();
        AltSourceSpan = image.AltSyntaxSpan;
        SourceSourceSpan = image.SourceSyntaxSpan;
        TitleSourceSpan = image.TitleSyntaxSpan;
        LinkUrlSourceSpan = image.LinkTargetSyntaxSpan;
        LinkTitleSourceSpan = image.LinkTitleSyntaxSpan;
    }

    /// <summary>Source image block.</summary>
    public ImageBlock Image { get; }

    /// <summary>Image source path or URL.</summary>
    public string Source { get; }

    /// <summary>Alternate text markdown when available.</summary>
    public string? Alt { get; }

    /// <summary>Plain alternate text when available.</summary>
    public string? PlainAlt { get; }

    /// <summary>Image title when available.</summary>
    public string? Title { get; }

    /// <summary>Requested display width when available.</summary>
    public double? Width { get; }

    /// <summary>Requested display height when available.</summary>
    public double? Height { get; }

    /// <summary>Optional image caption.</summary>
    public string? Caption { get; }

    /// <summary>Optional link target wrapping the image.</summary>
    public string? LinkUrl { get; }

    /// <summary>Optional link title.</summary>
    public string? LinkTitle { get; }

    /// <summary>Optional link HTML target.</summary>
    public string? LinkTarget { get; }

    /// <summary>Optional link rel value.</summary>
    public string? LinkRel { get; }

    /// <summary>Optional fallback path for picture sources.</summary>
    public string? PictureFallbackPath { get; }

    /// <summary>Responsive picture sources in source order.</summary>
    public IReadOnlyList<ImagePictureSource> PictureSources { get; }

    /// <summary>Source span for the image alternate text token when parsed from markdown.</summary>
    public MarkdownSourceSpan? AltSourceSpan { get; }

    /// <summary>Source span for the image source token when parsed from markdown.</summary>
    public MarkdownSourceSpan? SourceSourceSpan { get; }

    /// <summary>Source span for the image title token when parsed from markdown.</summary>
    public MarkdownSourceSpan? TitleSourceSpan { get; }

    /// <summary>Source span for the wrapping link target token when parsed from markdown.</summary>
    public MarkdownSourceSpan? LinkUrlSourceSpan { get; }

    /// <summary>Source span for the wrapping link title token when parsed from markdown.</summary>
    public MarkdownSourceSpan? LinkTitleSourceSpan { get; }
}

/// <summary>
/// Native projection for front matter.
/// </summary>
public sealed class MarkdownNativeFrontMatterBlock : MarkdownNativeBlock {
    internal MarkdownNativeFrontMatterBlock(FrontMatterBlock frontMatter, MarkdownSyntaxNode syntaxNode)
        : base(MarkdownNativeBlockKind.FrontMatter, frontMatter, syntaxNode) {
        FrontMatter = frontMatter;
        Entries = frontMatter.Entries;
        Values = frontMatter.Entries.ToDictionary(
            static entry => entry.Key,
            static entry => entry.Value,
            StringComparer.OrdinalIgnoreCase);
        OpeningFenceSourceSpan = GetChildSpan(syntaxNode, MarkdownSyntaxKind.FrontMatterOpeningFence);
        ClosingFenceSourceSpan = GetChildSpan(syntaxNode, MarkdownSyntaxKind.FrontMatterClosingFence);
    }

    /// <summary>Source front matter block.</summary>
    public FrontMatterBlock FrontMatter { get; }

    /// <summary>Structured front matter entries in source order.</summary>
    public IReadOnlyList<FrontMatterBlock.Entry> Entries { get; }

    /// <summary>Front matter values by key.</summary>
    public IReadOnlyDictionary<string, object?> Values { get; }

    /// <summary>Source span for the opening front matter fence marker when parsed from markdown.</summary>
    public MarkdownSourceSpan? OpeningFenceSourceSpan { get; }

    /// <summary>Source span for the closing front matter fence marker when parsed from markdown.</summary>
    public MarkdownSourceSpan? ClosingFenceSourceSpan { get; }

    private static MarkdownSourceSpan? GetChildSpan(MarkdownSyntaxNode syntaxNode, MarkdownSyntaxKind kind) =>
        syntaxNode?.Children.FirstOrDefault(child => child.Kind == kind)?.SourceSpan;
}

/// <summary>
/// Native projection for raw HTML and HTML comments.
/// </summary>
public sealed class MarkdownNativeHtmlBlock : MarkdownNativeBlock {
    internal MarkdownNativeHtmlBlock(HtmlRawBlock html, MarkdownSyntaxNode syntaxNode)
        : base(MarkdownNativeBlockKind.Html, html, syntaxNode) {
        Html = html.Html;
        IsComment = false;
    }

    internal MarkdownNativeHtmlBlock(HtmlCommentBlock comment, MarkdownSyntaxNode syntaxNode)
        : base(MarkdownNativeBlockKind.Html, comment, syntaxNode) {
        Html = comment.Comment;
        IsComment = true;
    }

    /// <summary>Raw HTML or comment text.</summary>
    public string Html { get; }

    /// <summary>Whether this block came from an HTML comment.</summary>
    public bool IsComment { get; }
}
