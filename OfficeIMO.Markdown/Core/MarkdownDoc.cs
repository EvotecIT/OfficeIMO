namespace OfficeIMO.Markdown;

/// <summary>
/// Root document container and fluent API entrypoint for composing Markdown.
/// Supports a fluent-chaining style and an explicit object model via <see cref="Add(IMarkdownBlock)"/>.
/// </summary>
public class MarkdownDoc {
    /// <summary>Resolved heading metadata within a document.</summary>
    public sealed class HeadingInfo {
        /// <summary>The heading block.</summary>
        public HeadingBlock Block { get; }
        /// <summary>Heading level.</summary>
        public int Level => Block.Level;
        /// <summary>Plain-text heading text.</summary>
        public string Text => Block.Text;
        /// <summary>Resolved anchor id for the heading within this document.</summary>
        public string Anchor { get; }

        internal HeadingInfo(HeadingBlock block, string anchor) {
            Block = block ?? throw new ArgumentNullException(nameof(block));
            Anchor = anchor ?? string.Empty;
        }
    }

    private readonly List<IMarkdownBlock> _blocks = new();
    private IMarkdownBlock? _lastBlock;
    private IFrontMatterMarkdownBlock? _frontMatter;

    /// <summary>Creates a new, empty Markdown document.</summary>
    public static MarkdownDoc Create() => new MarkdownDoc();

    /// <summary>All blocks added to the document (excluding front matter).</summary>
    public IReadOnlyList<IMarkdownBlock> Blocks => _blocks;
    /// <summary>Document-level front matter/header block when present.</summary>
    public FrontMatterBlock? DocumentHeader => _frontMatter as FrontMatterBlock;
    /// <summary>All top-level document blocks in order, including front matter when present.</summary>
    public IReadOnlyList<IMarkdownBlock> TopLevelBlocks => BuildTopLevelBlocks();
    /// <summary>Whether the document has front matter.</summary>
    public bool HasDocumentHeader => DocumentHeader != null;

    /// <summary>Adds a block instance (object-model style).</summary>
    /// <param name="block">Block to append to the document.</param>
    /// <returns>Same <see cref="MarkdownDoc"/> for chaining.</returns>
    public MarkdownDoc Add(IMarkdownBlock block) {
        if (block is IFrontMatterMarkdownBlock fm) {
            _frontMatter = fm;
        } else {
            _blocks.Add(block);
            _lastBlock = block;
        }
        return this;
    }

    /// <summary>Enumerates all document blocks depth-first, including front matter when present.</summary>
    public IEnumerable<IMarkdownBlock> DescendantsAndSelf() {
        foreach (var block in TopLevelBlocks) {
            foreach (var descendant in EnumerateBlockAndDescendants(block)) {
                yield return descendant;
            }
        }
    }

    /// <summary>Enumerates top-level document blocks of the requested type.</summary>
    public IEnumerable<TBlock> TopLevelBlocksOfType<TBlock>() where TBlock : class, IMarkdownBlock {
        foreach (var block in TopLevelBlocks) {
            if (block is TBlock typedBlock) {
                yield return typedBlock;
            }
        }
    }

    /// <summary>Finds the first top-level document block of the requested type.</summary>
    public TBlock? FindFirstTopLevelBlockOfType<TBlock>() where TBlock : class, IMarkdownBlock {
        foreach (var block in TopLevelBlocks) {
            if (block is TBlock typedBlock) {
                return typedBlock;
            }
        }

        return null;
    }

    /// <summary>Checks whether the document has a top-level block of the requested type.</summary>
    public bool HasTopLevelBlockOfType<TBlock>() where TBlock : class, IMarkdownBlock =>
        FindFirstTopLevelBlockOfType<TBlock>() != null;

    /// <summary>Enumerates all document blocks of the requested type depth-first.</summary>
    public IEnumerable<TBlock> DescendantsOfType<TBlock>() where TBlock : class, IMarkdownBlock {
        foreach (var block in DescendantsAndSelf()) {
            if (block is TBlock typedBlock) {
                yield return typedBlock;
            }
        }
    }

    /// <summary>Finds the first document block of the requested type depth-first.</summary>
    public TBlock? FindFirstDescendantOfType<TBlock>() where TBlock : class, IMarkdownBlock {
        foreach (var block in DescendantsAndSelf()) {
            if (block is TBlock typedBlock) {
                return typedBlock;
            }
        }

        return null;
    }

    /// <summary>Checks whether the document has any block of the requested type depth-first.</summary>
    public bool HasDescendantOfType<TBlock>() where TBlock : class, IMarkdownBlock =>
        FindFirstDescendantOfType<TBlock>() != null;

    /// <summary>Enumerates all list items in document order, including nested items.</summary>
    public IEnumerable<ListItem> DescendantListItems() {
        foreach (var block in TopLevelBlocks) {
            foreach (var item in EnumerateListItems(block)) {
                yield return item;
            }
        }
    }

    /// <summary>Enumerates all headings in document order, including nested headings.</summary>
    public IEnumerable<HeadingBlock> DescendantHeadings() {
        foreach (var block in DescendantsAndSelf()) {
            if (block is HeadingBlock heading) {
                yield return heading;
            }
        }
    }

    /// <summary>Gets the resolved anchor id for a heading within this document.</summary>
    public string GetHeadingAnchor(HeadingBlock heading) {
        if (heading == null) throw new ArgumentNullException(nameof(heading));

        var (_, headingCatalog) = GetBlocksAndHeadingSlugs();
        return headingCatalog.GetHeadingAnchor(heading);
    }

    /// <summary>Returns resolved heading metadata in document order.</summary>
    public IReadOnlyList<HeadingInfo> GetHeadingInfos() {
        var headings = DescendantHeadings().ToArray();
        if (headings.Length == 0) {
            return Array.Empty<HeadingInfo>();
        }

        var infos = new HeadingInfo[headings.Length];
        for (int i = 0; i < headings.Length; i++) {
            infos[i] = new HeadingInfo(headings[i], GetHeadingAnchor(headings[i]));
        }
        return infos;
    }

    /// <summary>Finds the heading with the specified resolved anchor, if present.</summary>
    public HeadingInfo? FindHeadingByAnchor(string anchor) {
        if (string.IsNullOrWhiteSpace(anchor)) {
            return null;
        }

        var normalized = anchor.Trim();
        if (normalized.StartsWith("#", StringComparison.Ordinal)) {
            normalized = normalized.Substring(1);
        }

        var headings = GetHeadingInfos();
        for (int i = 0; i < headings.Count; i++) {
            if (string.Equals(headings[i].Anchor, normalized, StringComparison.Ordinal)) {
                return headings[i];
            }
        }

        return null;
    }

    /// <summary>Checks whether a heading with the specified resolved anchor is present.</summary>
    public bool HasHeadingAnchor(string anchor) => FindHeadingByAnchor(anchor) != null;

    /// <summary>Finds the first heading whose plain text matches the provided heading text.</summary>
    public HeadingInfo? FindHeading(string text, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
        if (string.IsNullOrEmpty(text)) {
            return null;
        }

        var headings = GetHeadingInfos();
        for (int i = 0; i < headings.Count; i++) {
            if (string.Equals(headings[i].Text, text, comparison)) {
                return headings[i];
            }
        }

        return null;
    }

    /// <summary>Checks whether a heading whose plain text matches the provided heading text is present.</summary>
    public bool HasHeading(string text, StringComparison comparison = StringComparison.OrdinalIgnoreCase) =>
        FindHeading(text, comparison) != null;

    /// <summary>Finds headings whose plain text matches the provided heading text.</summary>
    public IReadOnlyList<HeadingInfo> FindHeadings(string text, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
        if (string.IsNullOrEmpty(text)) {
            return Array.Empty<HeadingInfo>();
        }

        var headings = GetHeadingInfos();
        var matches = new List<HeadingInfo>();
        for (int i = 0; i < headings.Count; i++) {
            if (string.Equals(headings[i].Text, text, comparison)) {
                matches.Add(headings[i]);
            }
        }

        return matches;
    }

    /// <summary>Finds a front matter entry by key when the document header is present.</summary>
    public FrontMatterBlock.Entry? FindFrontMatterEntry(string key, StringComparison comparison = StringComparison.OrdinalIgnoreCase) =>
        DocumentHeader?.FindEntry(key, comparison);

    /// <summary>Checks whether the document header contains an entry with the specified key.</summary>
    public bool HasFrontMatterEntry(string key, StringComparison comparison = StringComparison.OrdinalIgnoreCase) =>
        DocumentHeader?.HasEntry(key, comparison) == true;

    /// <summary>Gets a typed front matter value by key when the document header is present.</summary>
    public bool TryGetFrontMatterValue<T>(string key, out T? value) {
        if (DocumentHeader != null && DocumentHeader.TryGetValue<T>(key, out value)) {
            return true;
        }

        value = default;
        return false;
    }

    private IReadOnlyList<IMarkdownBlock> BuildTopLevelBlocks() {
        if (_frontMatter == null) {
            return _blocks;
        }

        var blocks = new List<IMarkdownBlock>(_blocks.Count + 1) {
            (IMarkdownBlock)_frontMatter
        };
        blocks.AddRange(_blocks);
        return blocks;
    }

    private static IEnumerable<IMarkdownBlock> EnumerateBlockAndDescendants(IMarkdownBlock block) {
        yield return block;

        if (block is IMarkdownListBlock listBlock) {
            for (int i = 0; i < listBlock.ListItems.Count; i++) {
                var item = listBlock.ListItems[i];
                for (int j = 0; j < item.BlockChildren.Count; j++) {
                    foreach (var descendant in EnumerateBlockAndDescendants(item.BlockChildren[j])) {
                        yield return descendant;
                    }
                }
            }

            yield break;
        }

        if (block is IChildMarkdownBlockContainer container) {
            for (int i = 0; i < container.ChildBlocks.Count; i++) {
                foreach (var descendant in EnumerateBlockAndDescendants(container.ChildBlocks[i])) {
                    yield return descendant;
                }
            }
        }
    }

    private static IEnumerable<ListItem> EnumerateListItems(IMarkdownBlock block) {
        if (block is IMarkdownListBlock listBlock) {
            for (int i = 0; i < listBlock.ListItems.Count; i++) {
                var item = listBlock.ListItems[i];
                yield return item;

                for (int j = 0; j < item.BlockChildren.Count; j++) {
                    foreach (var descendant in EnumerateListItems(item.BlockChildren[j])) {
                        yield return descendant;
                    }
                }
            }

            yield break;
        }

        if (block is IChildMarkdownBlockContainer container) {
            for (int i = 0; i < container.ChildBlocks.Count; i++) {
                foreach (var descendant in EnumerateListItems(container.ChildBlocks[i])) {
                    yield return descendant;
                }
            }
        }
    }

    /// <summary>Sets YAML front matter from an anonymous object or dictionary.</summary>
    public MarkdownDoc FrontMatter(object data) {
        _frontMatter = FrontMatterBlock.FromObject(data);
        return this;
    }

    /// <summary>Adds an H1 heading.</summary>
    public MarkdownDoc H1(string text) => Add(new HeadingBlock(1, text));
    /// <summary>Adds an H2 heading.</summary>
    public MarkdownDoc H2(string text) => Add(new HeadingBlock(2, text));
    /// <summary>Adds an H3 heading.</summary>
    public MarkdownDoc H3(string text) => Add(new HeadingBlock(3, text));
    /// <summary>Adds an H4 heading.</summary>
    public MarkdownDoc H4(string text) => Add(new HeadingBlock(4, text));
    /// <summary>Adds an H5 heading.</summary>
    public MarkdownDoc H5(string text) => Add(new HeadingBlock(5, text));
    /// <summary>Adds an H6 heading.</summary>
    public MarkdownDoc H6(string text) => Add(new HeadingBlock(6, text));

    /// <summary>Adds a paragraph with plain text.</summary>
    public MarkdownDoc P(string text) => Add(new ParagraphBlock(new InlineSequence().Text(text)));

    /// <summary>Adds a paragraph composed with the paragraph builder.</summary>
    public MarkdownDoc P(Action<ParagraphBuilder> build) {
        ParagraphBuilder builder = new ParagraphBuilder();
        build(builder);
        return Add(new ParagraphBlock(builder.Inlines));
    }

    /// <summary>Adds a simple block quote with a single line of text.</summary>
    public MarkdownDoc Quote(string text) => Add(new QuoteBlock(new[] { text ?? string.Empty }));

    /// <summary>Adds a block quote composed via <see cref="QuoteBuilder"/>.</summary>
    public MarkdownDoc Quote(Action<QuoteBuilder> build) {
        QuoteBuilder builder = new QuoteBuilder();
        build(builder);
        return Add(builder.Build());
    }

    /// <summary>Adds a collapsible details block.</summary>
    public MarkdownDoc Details(string summary, Action<MarkdownDoc> buildBody, bool open = false) {
        var inner = MarkdownDoc.Create();
        buildBody(inner);
        var details = new DetailsBlock(new SummaryBlock(summary), inner.Blocks, open);
        return Add(details);
    }

    /// <summary>Adds a horizontal rule.</summary>
    public MarkdownDoc Hr() => Add(new HorizontalRuleBlock());

    /// <summary>Adds a callout/admonition block (Docs-style).</summary>
    public MarkdownDoc Callout(string kind, string title, string body) => Add(new CalloutBlock(kind, title, body));

    /// <summary>Adds an unordered list.</summary>
    public MarkdownDoc Ul(Action<UnorderedListBuilder> build) {
        UnorderedListBuilder builder = new UnorderedListBuilder();
        build(builder);
        return Add(builder.Build());
    }

    /// <summary>Adds an ordered list.</summary>
    public MarkdownDoc Ol(Action<OrderedListBuilder> build) {
        OrderedListBuilder builder = new OrderedListBuilder();
        build(builder);
        return Add(builder.Build());
    }

    /// <summary>Adds a definition list.</summary>
    public MarkdownDoc Dl(Action<DefinitionListBuilder> build) {
        DefinitionListBuilder builder = new DefinitionListBuilder();
        build(builder);
        return Add(builder.Build());
    }

    /// <summary>Adds a fenced code block.</summary>
    public MarkdownDoc Code(string language, string content) {
        CodeBlock code = new CodeBlock(language, content);
        Add(code);
        return this;
    }

    /// <summary>Sets a caption for the last captionable block (image/code), or appends a paragraph.</summary>
    public MarkdownDoc Caption(string caption) {
        if (_lastBlock is ICaptionable cap) {
            cap.Caption = caption;
        } else {
            // Fallback: render as simple paragraph text
            P(caption);
        }
        return this;
    }

    /// <summary>Adds an image block with optional alt text and title.</summary>
    public MarkdownDoc Image(string path, string? alt, string? title)
        => Add(new ImageBlock(path, alt, title));

    /// <summary>Adds an image block with optional alt text, title, and size hints.</summary>
    public MarkdownDoc Image(string path, string? alt = null, string? title = null, double? width = null, double? height = null)
        => Add(new ImageBlock(path, alt, title, width, height));

    /// <summary>Adds a table built with <see cref="TableBuilder"/>.</summary>
    public MarkdownDoc Table(Action<TableBuilder> build) {
        TableBuilder tb = new TableBuilder();
        build(tb);
        return Add(tb.Build());
    }

    /// <summary>
    /// Builds a table via <see cref="TableBuilder"/>, then applies auto-alignment heuristics if requested.
    /// </summary>
    public MarkdownDoc TableAuto(Action<TableBuilder> build, bool alignNumeric = true, bool alignDates = true) {
        TableBuilder tb = new TableBuilder();
        build(tb);
        if (alignDates) tb.AlignDatesCenter();
        if (alignNumeric) tb.AlignNumericRight();
        return Add(tb.Build());
    }

    /// <summary>
    /// Convenience to add a table from arbitrary data; see <see cref="TableBuilder.FromAny(object?)"/> for rules.
    /// </summary>
    public MarkdownDoc TableFrom(object? data) {
        TableBuilder tb = new TableBuilder();
        tb.FromAny(data);
        return Add(tb.Build());
    }

    /// <summary>
    /// Creates a table from data and applies auto-alignment heuristics (numeric right, dates center) if requested.
    /// </summary>
    public MarkdownDoc TableFromAuto(object? data, System.Action<TableFromOptions>? configure = null, bool alignNumeric = true, bool alignDates = true) {
        TableBuilder tb = new TableBuilder();
        if (configure is null) tb.FromAny(data); else tb.FromAny(data, configure);
        if (alignDates) tb.AlignDatesCenter();
        if (alignNumeric) tb.AlignNumericRight();
        return Add(tb.Build());
    }

    /// <summary>
    /// Convenience to add a table from a sequence using column selectors.
    /// </summary>
    public MarkdownDoc TableFrom<T>(System.Collections.Generic.IEnumerable<T> items, params (string Header, System.Func<T, object?> Selector)[] columns) {
        TableBuilder tb = new TableBuilder();
        tb.FromSequence(items, columns);
        return Add(tb.Build());
    }

    /// <summary>Adds an unordered list from a sequence of items using ToString().</summary>
    public MarkdownDoc Ul<T>(System.Collections.Generic.IEnumerable<T> items) {
        UnorderedListBuilder builder = new UnorderedListBuilder();
        builder.Items(items, null);
        return Add(builder.Build());
    }

    /// <summary>Adds an ordered list from a sequence of items using ToString().</summary>
    public MarkdownDoc Ol<T>(System.Collections.Generic.IEnumerable<T> items, int start = 1) {
        OrderedListBuilder builder = new OrderedListBuilder().StartAt(start);
        builder.Items(items, null);
        return Add(builder.Build());
    }

    /// <summary>Renders the document to Markdown string.</summary>
    public string ToMarkdown() {
        // Build a transient block list where TOC placeholders are realized
        var (blocks, _) = GetBlocksAndHeadingSlugs();
        StringBuilder sb = new StringBuilder();
        if (_frontMatter != null) {
            sb.AppendLine(_frontMatter.RenderFrontMatter());
            sb.AppendLine();
        }
        for (int i = 0; i < blocks.Count; i++) {
            string rendered = blocks[i].RenderMarkdown();
            if (!string.IsNullOrEmpty(rendered)) sb.AppendLine(rendered);
            if (i < blocks.Count - 1) sb.AppendLine();
        }
        return sb.ToString();
    }

    /// <summary>
    /// Renders HTML using default options. For backward compatibility, this returns an embeddable HTML fragment
    /// (no html/head/body) containing just the rendered content. Use <see cref="ToHtmlDocument"/> for a full page.
    /// </summary>
    public string ToHtml() => ToHtmlFragment();

    /// <summary>Renders an embeddable HTML fragment. Wraps in &lt;article class="markdown-body"&gt; by default.</summary>
    public string ToHtmlFragment(HtmlOptions? options = null) {
        options ??= new HtmlOptions { Kind = HtmlKind.Fragment };
        options.Kind = HtmlKind.Fragment;
        return HtmlRenderer.Render(this, options);
    }

    /// <summary>
    /// Renders an embeddable HTML fragment and inlines the computed CSS and tiny scripts at the top/bottom
    /// of the fragment. Useful when you want a single self-contained chunk without a full HTML document.
    /// </summary>
    public string ToHtmlFragmentWithCss(HtmlOptions? options = null) {
        options ??= new HtmlOptions { Kind = HtmlKind.Fragment };
        options.Kind = HtmlKind.Fragment;
        var parts = HtmlRenderer.RenderParts(this, options);
        var sb = new StringBuilder();
        if (!string.IsNullOrEmpty(parts.Css)) sb.Append("<style>\n").Append(parts.Css).Append("\n</style>");
        sb.Append(parts.Body);
        if (!string.IsNullOrEmpty(parts.Scripts)) sb.Append("<script>\n").Append(parts.Scripts).Append("\n</script>");
        return sb.ToString();
    }

    /// <summary>Renders a standalone HTML5 document with optional CSS/JS assets.</summary>
    public string ToHtmlDocument(HtmlOptions? options = null) {
        options ??= new HtmlOptions { Kind = HtmlKind.Document };
        options.Kind = HtmlKind.Document;
        return HtmlRenderer.Render(this, options);
    }

    /// <summary>Asynchronously renders an embeddable HTML fragment.</summary>
    public System.Threading.Tasks.Task<string> ToHtmlFragmentAsync(HtmlOptions? options = null) => System.Threading.Tasks.Task.FromResult(ToHtmlFragment(options));
    /// <summary>Asynchronously renders a full HTML document.</summary>
    public System.Threading.Tasks.Task<string> ToHtmlDocumentAsync(HtmlOptions? options = null) => System.Threading.Tasks.Task.FromResult(ToHtmlDocument(options));

    /// <summary>Returns rendered parts for advanced embedding (Head, Body, Css, Scripts).</summary>
    public HtmlRenderParts ToHtmlParts(HtmlOptions? options = null) {
        options ??= new HtmlOptions { Kind = HtmlKind.Fragment };
        return HtmlRenderer.RenderParts(this, options);
    }

    internal (System.Collections.Generic.List<IMarkdownBlock> Blocks, MarkdownHeadingCatalog HeadingCatalog) GetBlocksAndHeadingSlugs() {
        var registry = MarkdownSlug.CreateRegistry();
        var (realized, headingCatalog) = RealizeTocPlaceholders(registry);
        return (realized, headingCatalog);
    }

    /// <summary>
    /// Saves HTML to the specified file. When <see cref="CssDelivery.ExternalFile"/> is used,
    /// writes a sidecar CSS file next to the HTML and links it.
    /// </summary>
    public void SaveHtml(string path, HtmlOptions? options = null) {
        options ??= new HtmlOptions();
        // If external CSS requested, compute sidecar path and let renderer know
        if (options.CssDelivery == CssDelivery.ExternalFile) {
            var basePath = System.IO.Path.ChangeExtension(path, null);
            var cssPath = basePath + ".css";
            options.ExternalCssOutputPath = cssPath;
        }
        var html = options.Kind == HtmlKind.Document ? ToHtmlDocument(options) : ToHtmlFragment(options);
        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(path)) ?? ".");
        System.IO.File.WriteAllText(path, html, System.Text.Encoding.UTF8);
        // If renderer produced a sidecar css, ensure it's written
        if (!string.IsNullOrEmpty(options.ExternalCssOutputPath) && options._externalCssContentToWrite is not null) {
            System.IO.File.WriteAllText(options.ExternalCssOutputPath!, options._externalCssContentToWrite, System.Text.Encoding.UTF8);
        }
    }

    /// <summary>
    /// Asynchronously saves HTML to the specified file. When <see cref="CssDelivery.ExternalFile"/> is used,
    /// writes a sidecar CSS file next to the HTML and links it.
    /// </summary>
    public async System.Threading.Tasks.Task SaveHtmlAsync(string path, HtmlOptions? options = null) {
        options ??= new HtmlOptions();
        if (options.CssDelivery == CssDelivery.ExternalFile) {
            var basePath = System.IO.Path.ChangeExtension(path, null);
            var cssPath = basePath + ".css";
            options.ExternalCssOutputPath = cssPath;
        }
        var html = options.Kind == HtmlKind.Document ? ToHtmlDocument(options) : ToHtmlFragment(options);
        System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(path)) ?? ".");
        await FileCompat.WriteAllTextAsync(path, html, System.Text.Encoding.UTF8).ConfigureAwait(false);
        if (!string.IsNullOrEmpty(options.ExternalCssOutputPath) && options._externalCssContentToWrite is not null) {
            await FileCompat.WriteAllTextAsync(options.ExternalCssOutputPath!, options._externalCssContentToWrite, System.Text.Encoding.UTF8).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Generates a Table of Contents from headings already present in the document and inserts it.
    /// </summary>
    /// <param name="configure">Optional TOC options.</param>
    /// <param name="placeAtTop">When true, inserts TOC as the first block; otherwise appended.</param>
    public MarkdownDoc Toc(System.Action<TocOptions>? configure = null, bool placeAtTop = false) {
        var opts = new TocOptions();
        configure?.Invoke(opts);

        var placeholder = new TocPlaceholderBlock(opts);
        if (opts.IncludeTitle) {
            // Insert a title heading above the TOC
            var heading = new HeadingBlock(opts.TitleLevel, opts.Title);
            if (placeAtTop) {
                _blocks.Insert(0, heading);
                _blocks.Insert(1, placeholder);
            } else {
                _blocks.Add(heading);
                _blocks.Add(placeholder);
            }
        } else {
            if (placeAtTop) _blocks.Insert(0, placeholder); else _blocks.Add(placeholder);
        }
        _lastBlock = placeholder;
        return this;
    }

    /// <summary>
    /// Convenience helper to insert a Table of Contents at the top with common parameters.
    /// </summary>
    public MarkdownDoc TocAtTop(string title = "Contents", int min = 1, int max = 3, bool ordered = false, int titleLevel = 2) {
        return Toc(opts => { opts.Title = title; opts.MinLevel = min; opts.MaxLevel = max; opts.Ordered = ordered; opts.TitleLevel = titleLevel; }, placeAtTop: true);
    }

    /// <summary>
    /// Inserts a TOC placeholder at the current position without a title heading by default.
    /// </summary>
    public MarkdownDoc TocHere(System.Action<TocOptions>? configure = null) {
        var opts = new TocOptions { IncludeTitle = false };
        configure?.Invoke(opts);
        var placeholder = new TocPlaceholderBlock(opts);
        _blocks.Add(placeholder);
        _lastBlock = placeholder;
        return this;
    }

    /// <summary>
    /// Inserts a section TOC for the nearest preceding heading. Useful to place a small TOC under a section.
    /// </summary>
    public MarkdownDoc TocForPreviousHeading(string? title = "Contents", int min = 2, int max = 6, bool ordered = false, int titleLevel = 3) {
        return Toc(opts => { opts.Title = title ?? ""; opts.IncludeTitle = !string.IsNullOrEmpty(title); opts.MinLevel = min; opts.MaxLevel = max; opts.Ordered = ordered; opts.TitleLevel = titleLevel; opts.Scope = TocScope.PreviousHeading; }, placeAtTop: false);
    }

    /// <summary>
    /// Inserts a section TOC scoped to the named heading.
    /// </summary>
    public MarkdownDoc TocForSection(string headingTitle, string? title = "Contents", int min = 2, int max = 6, bool ordered = false, int titleLevel = 3) {
        return Toc(opts => { opts.Title = title ?? ""; opts.IncludeTitle = !string.IsNullOrEmpty(title); opts.MinLevel = min; opts.MaxLevel = max; opts.Ordered = ordered; opts.TitleLevel = titleLevel; opts.Scope = TocScope.HeadingTitle; opts.ScopeHeadingTitle = headingTitle; }, placeAtTop: false);
    }

    private (System.Collections.Generic.List<IMarkdownBlock> Blocks, MarkdownHeadingCatalog HeadingCatalog) RealizeTocPlaceholders(System.Collections.Generic.Dictionary<string, int> slugRegistry) {
        // Create a shallow copy first
        var realized = new System.Collections.Generic.List<IMarkdownBlock>(_blocks);
        var headingCatalog = MarkdownHeadingCatalog.Create(realized, slugRegistry);
        // Replace placeholders with generated TOC blocks
        for (int i = 0; i < realized.Count; i++) {
            if (realized[i] is ITocPlaceholderMarkdownBlock tocPlaceholder) {
                realized[i] = tocPlaceholder.RealizeToc(realized, i, headingCatalog);
            }
        }
        return (realized, headingCatalog);
    }
}
