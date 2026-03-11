namespace OfficeIMO.Markdown;

internal sealed class MarkdownHeadingCatalog {
    internal readonly struct HeadingEntry {
        internal HeadingEntry(int index, HeadingBlock block, string anchor) {
            Index = index;
            Block = block;
            Anchor = anchor ?? string.Empty;
        }

        internal int Index { get; }
        internal HeadingBlock Block { get; }
        internal int Level => Block.Level;
        internal string Text => Block.Text;
        internal string Anchor { get; }
    }

    private readonly IReadOnlyList<HeadingEntry> _headings;

    private MarkdownHeadingCatalog(IReadOnlyList<HeadingEntry> headings, IReadOnlyDictionary<HeadingBlock, string> headingSlugs) {
        _headings = headings;
        HeadingSlugs = headingSlugs;
    }

    internal IReadOnlyDictionary<HeadingBlock, string> HeadingSlugs { get; }

    internal static MarkdownHeadingCatalog Create(IReadOnlyList<IMarkdownBlock> blocks, Dictionary<string, int>? slugRegistry = null) {
        var headings = new List<HeadingEntry>();
        var slugs = new Dictionary<HeadingBlock, string>();

        for (int idx = 0; idx < blocks.Count; idx++) {
            if (blocks[idx] is not HeadingBlock heading) {
                continue;
            }

            var slug = slugRegistry == null
                ? MarkdownSlug.GitHub(heading.Text)
                : MarkdownSlug.GitHub(heading.Text, slugRegistry);

            slugs[heading] = slug;
            headings.Add(new HeadingEntry(idx, heading, slug));
        }

        return new MarkdownHeadingCatalog(headings, slugs);
    }

    internal string? GetPrecedingHeadingAnchor(IReadOnlyList<IMarkdownBlock> blocks, int blockIndex, TocOptions options) {
        if (!options.IncludeTitle || blockIndex <= 0 || blockIndex > blocks.Count - 1) {
            return null;
        }

        if (blocks[blockIndex - 1] is not HeadingBlock titleHeading) {
            return null;
        }

        if (!HeadingSlugs.TryGetValue(titleHeading, out var titleSlug)) {
            titleSlug = MarkdownSlug.GitHub(titleHeading.Text);
        }

        return titleSlug;
    }

    internal List<TocBlock.Entry> BuildTocEntries(IReadOnlyList<IMarkdownBlock> blocks, int placeholderIndex, TocOptions options, string? titleAnchor = null) {
        int minLevel = ClampHeadingLevel(options.MinLevel);
        int maxLevel = ClampHeadingLevel(options.MaxLevel);
        if (maxLevel < minLevel) {
            maxLevel = minLevel;
        }

        int effectiveMin = options.RequireTopLevel && minLevel > 1 ? 1 : minLevel;
        int effectiveMax = maxLevel;

        GetScopeBounds(placeholderIndex, minLevel, options, out var startIndex, out var endIndex);

        var entries = _headings
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

    private void GetScopeBounds(int placeholderIndex, int minLevel, TocOptions options, out int startIndex, out int endIndex) {
        startIndex = 0;
        endIndex = int.MaxValue;

        if (options.Scope == TocScope.PreviousHeading) {
            var previous = _headings.LastOrDefault(h => h.Index < placeholderIndex && h.Level < minLevel);
            if (previous.Equals(default(HeadingEntry))) {
                previous = _headings.LastOrDefault(h => h.Index < placeholderIndex);
            }

            if (!previous.Equals(default(HeadingEntry))) {
                startIndex = previous.Index + 1;
                var nextAtOrAbove = _headings.FirstOrDefault(h => h.Index > previous.Index && h.Level <= previous.Level);
                if (!nextAtOrAbove.Equals(default(HeadingEntry))) {
                    endIndex = nextAtOrAbove.Index;
                }
            }

            return;
        }

        if (options.Scope == TocScope.HeadingTitle && !string.IsNullOrWhiteSpace(options.ScopeHeadingTitle)) {
            var root = _headings.FirstOrDefault(h => string.Equals(h.Text, options.ScopeHeadingTitle, StringComparison.OrdinalIgnoreCase));
            if (!root.Equals(default(HeadingEntry))) {
                startIndex = root.Index + 1;
                var nextAtOrAbove = _headings.FirstOrDefault(h => h.Index > root.Index && h.Level <= root.Level);
                if (!nextAtOrAbove.Equals(default(HeadingEntry))) {
                    endIndex = nextAtOrAbove.Index;
                }
            }
        }
    }

    private static int ClampHeadingLevel(int level) => level < 1 ? 1 : (level > 6 ? 6 : level);
}
