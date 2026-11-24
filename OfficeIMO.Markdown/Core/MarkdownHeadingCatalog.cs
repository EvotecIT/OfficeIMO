namespace OfficeIMO.Markdown;

internal sealed class MarkdownHeadingCatalog {
    internal readonly struct HeadingEntry {
        internal HeadingEntry(int index, IHeadingMarkdownBlock block, string anchor) {
            Index = index;
            Block = block;
            Anchor = anchor ?? string.Empty;
        }

        internal int Index { get; }
        internal IHeadingMarkdownBlock Block { get; }
        internal int Level => Block.Level;
        internal string Text => Block.Text;
        internal string Anchor { get; }
    }

    // Anchors are resolved on first use: most renders never ask for one, and resolving them
    // creates a slug string per heading.
    private readonly IReadOnlyList<IMarkdownBlock> _blocks;
    private readonly MarkdownHeadingIdentifierStyle _style;
    private readonly System.Threading.CancellationToken _cancellationToken;
    private List<HeadingEntry>? _headingEntries;
    private Dictionary<IHeadingMarkdownBlock, string>? _headingSlugs;

    private MarkdownHeadingCatalog(
        IReadOnlyList<IMarkdownBlock> blocks,
        MarkdownHeadingIdentifierStyle style,
        System.Threading.CancellationToken cancellationToken) {
        _blocks = blocks;
        _style = style;
        _cancellationToken = cancellationToken;
    }

    private List<HeadingEntry> Headings {
        get {
            EnsureResolved();
            return _headingEntries!;
        }
    }

    internal IReadOnlyDictionary<IHeadingMarkdownBlock, string> HeadingSlugs {
        get {
            EnsureResolved();
            return _headingSlugs!;
        }
    }

    internal string GetHeadingAnchor(IHeadingMarkdownBlock heading) {
        if (heading == null) {
            return string.Empty;
        }

        if (heading is HeadingBlock { SuppressAutoIdentifier: true }) {
            return string.Empty;
        }

        if (HeadingSlugs.TryGetValue(heading, out var slug)) {
            return slug;
        }

        return MarkdownSlug.Generate(heading.Text, _style);
    }

    /// <summary>
    /// Creates the catalog for <paramref name="blocks"/>, giving repeated heading texts distinct anchors.
    /// Anchors are resolved the first time one is requested, so <paramref name="blocks"/> must not change before then.
    /// </summary>
    internal static MarkdownHeadingCatalog Create(
        IReadOnlyList<IMarkdownBlock> blocks,
        MarkdownHeadingIdentifierStyle style,
        System.Threading.CancellationToken cancellationToken) =>
        new MarkdownHeadingCatalog(blocks, style, cancellationToken);

    private void EnsureResolved() {
        if (_headingEntries != null) {
            return;
        }

        var slugRegistry = MarkdownSlug.CreateRegistry();
        var headings = new List<HeadingEntry>();
        var slugs = new Dictionary<IHeadingMarkdownBlock, string>();

        for (int idx = 0; idx < _blocks.Count; idx++) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (_blocks[idx] is not IHeadingMarkdownBlock heading) {
                continue;
            }

            var slug = ResolveHeadingAnchor(heading, slugRegistry, _style);

            slugs[heading] = slug;
            headings.Add(new HeadingEntry(idx, heading, slug));
        }
        _cancellationToken.ThrowIfCancellationRequested();

        _headingSlugs = slugs;
        _headingEntries = headings;
    }

    private static string ResolveHeadingAnchor(
        IHeadingMarkdownBlock heading,
        Dictionary<string, int>? slugRegistry,
        MarkdownHeadingIdentifierStyle style) {
        if (heading is HeadingBlock { SuppressAutoIdentifier: true }) {
            return string.Empty;
        }

        if (heading is MarkdownObject { Attributes.ElementId: { } explicitId }
            && !string.IsNullOrWhiteSpace(explicitId)) {
            var anchor = explicitId.Trim();
            ReserveSlug(anchor, slugRegistry);
            return anchor;
        }

        return slugRegistry == null
            ? MarkdownSlug.Generate(heading.Text, style)
            : MarkdownSlug.Generate(heading.Text, style, slugRegistry);
    }

    private static void ReserveSlug(string anchor, Dictionary<string, int>? slugRegistry) {
        if (slugRegistry == null || string.IsNullOrEmpty(anchor)) {
            return;
        }

        if (slugRegistry.TryGetValue(anchor, out var count)) {
            slugRegistry[anchor] = count + 1;
            return;
        }

        slugRegistry[anchor] = 0;
    }

    internal string? GetPrecedingHeadingAnchor(IReadOnlyList<IMarkdownBlock> blocks, int blockIndex, MarkdownTocOptions options) {
        if (!options.IncludeTitle || blockIndex <= 0 || blockIndex > blocks.Count - 1) {
            return null;
        }

        if (blocks[blockIndex - 1] is not IHeadingMarkdownBlock titleHeading) {
            return null;
        }

        return GetHeadingAnchor(titleHeading);
    }

    internal List<TocBlock.Entry> BuildTocEntries(IReadOnlyList<IMarkdownBlock> blocks, int placeholderIndex, MarkdownTocOptions options, string? titleAnchor = null) {
        int minLevel = ClampHeadingLevel(options.MinLevel);
        int maxLevel = ClampHeadingLevel(options.MaxLevel);
        if (maxLevel < minLevel) {
            maxLevel = minLevel;
        }

        int effectiveMin = options.RequireTopLevel && minLevel > 1 ? 1 : minLevel;
        int effectiveMax = maxLevel;

        GetScopeBounds(placeholderIndex, minLevel, options, out var startIndex, out var endIndex);

        var entries = Headings
            .Where(h => h.Index >= startIndex && h.Index < endIndex && h.Level >= effectiveMin && h.Level <= effectiveMax)
            .Select(h => new TocBlock.Entry { Level = h.Level, Text = h.Text, Anchor = h.Anchor })
            .ToList();

        if (!string.IsNullOrEmpty(titleAnchor)) {
            entries = entries
                .Where(e => !string.Equals(e.Anchor, titleAnchor, StringComparison.Ordinal))
                .ToList();
        }

        return entries;
    }

    private void GetScopeBounds(int placeholderIndex, int minLevel, MarkdownTocOptions options, out int startIndex, out int endIndex) {
        startIndex = 0;
        endIndex = int.MaxValue;

        if (options.Scope == MarkdownTocScope.PreviousHeading) {
            var previous = Headings.LastOrDefault(h => h.Index < placeholderIndex && h.Level < minLevel);
            if (previous.Equals(default(HeadingEntry))) {
                previous = Headings.LastOrDefault(h => h.Index < placeholderIndex);
            }

            if (!previous.Equals(default(HeadingEntry))) {
                startIndex = previous.Index + 1;
                var nextAtOrAbove = Headings.FirstOrDefault(h => h.Index > previous.Index && h.Level <= previous.Level);
                if (!nextAtOrAbove.Equals(default(HeadingEntry))) {
                    endIndex = nextAtOrAbove.Index;
                }
            }

            return;
        }

        if (options.Scope == MarkdownTocScope.HeadingTitle && !string.IsNullOrWhiteSpace(options.ScopeHeadingTitle)) {
            var root = Headings.FirstOrDefault(h => string.Equals(h.Text, options.ScopeHeadingTitle, StringComparison.OrdinalIgnoreCase));
            if (!root.Equals(default(HeadingEntry))) {
                startIndex = root.Index + 1;
                var nextAtOrAbove = Headings.FirstOrDefault(h => h.Index > root.Index && h.Level <= root.Level);
                if (!nextAtOrAbove.Equals(default(HeadingEntry))) {
                    endIndex = nextAtOrAbove.Index;
                }
            }
        }
    }

    private static int ClampHeadingLevel(int level) => level < 1 ? 1 : (level > 6 ? 6 : level);
}
