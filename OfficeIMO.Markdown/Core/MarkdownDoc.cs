using OfficeIMO.Core.Internal;

namespace OfficeIMO.Markdown;

/// <summary>
/// Root document container for Markdown blocks and syntax helpers.
/// Use <see cref="Add{T}(T)"/> to append and retain a typed block.
/// </summary>
public partial class MarkdownDocument : MarkdownObject {
    internal MarkdownDocument() { }
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
    private MarkdownParseInfo? _parseResult;

    /// <summary>Creates a new, empty Markdown document.</summary>
    public static MarkdownDocument Create(MarkdownCreateOptions? options = null) {
        MarkdownDocument document = options?.Template == null ? new MarkdownDocument() : Load(options.Template);
        document._filePath = null;
        return document;
    }

    /// <summary>Creates a Markdown document associated with a path, without writing the file.</summary>
    public static MarkdownDocument Create(string filePath, MarkdownCreateOptions? options = null) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
        MarkdownDocument document = Create(options);
        document._filePath = System.IO.Path.GetFullPath(filePath);
        return document;
    }

    /// <summary>Rewrites the document using the provided rewriter.</summary>
    public MarkdownDocument Rewrite(MarkdownRewriter rewriter) {
        if (rewriter == null) {
            throw new ArgumentNullException(nameof(rewriter));
        }

        return rewriter.Rewrite(this);
    }

    internal void ReplaceBlocks(IEnumerable<IMarkdownBlock>? blocks) {
        _blocks.Clear();
        _lastBlock = null;
        _parseResult = null;
        if (blocks == null) {
            MarkdownObjectTreeBinder.BindDocument(this);
            return;
        }

        foreach (var block in blocks) {
            if (block == null) {
                continue;
            }

            _blocks.Add(block);
            _lastBlock = block;
        }

        MarkdownObjectTreeBinder.BindDocument(this);
    }

    internal void AttachParseResult(MarkdownParseInfo parseResult) {
        _parseResult = parseResult ?? throw new ArgumentNullException(nameof(parseResult));
    }

    /// <summary>Enumerates all document blocks depth-first, including front matter when present.</summary>
    public IEnumerable<IMarkdownBlock> DescendantsAndSelf() {
        MarkdownObjectTreeBinder.BindDocument(this);
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
        MarkdownObjectTreeBinder.BindDocument(this);
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

    /// <summary>Enumerates all tables in document order, including nested tables.</summary>
    public IEnumerable<TableBlock> DescendantTables() {
        foreach (var block in DescendantsAndSelf()) {
            if (block is TableBlock table) {
                yield return table;
            }
        }
    }

    /// <summary>Enumerates all table rows in document order, including header rows and nested tables.</summary>
    public IEnumerable<TableRow> DescendantTableRows() {
        foreach (var table in DescendantTables()) {
            foreach (var row in table.EnumerateRows()) {
                yield return row;
            }
        }
    }

    /// <summary>Enumerates all table cells in document order, including header cells and nested tables.</summary>
    public IEnumerable<TableCell> DescendantTableCells() {
        foreach (var table in DescendantTables()) {
            foreach (var cell in table.EnumerateCells()) {
                yield return cell;
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
        MarkdownObjectTreeBinder.BindDocument(this);
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
                for (int j = 0; j < item.ChildBlocks.Count; j++) {
                    foreach (var descendant in EnumerateBlockAndDescendants(item.ChildBlocks[j])) {
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

                for (int j = 0; j < item.ChildBlocks.Count; j++) {
                    foreach (var descendant in EnumerateListItems(item.ChildBlocks[j])) {
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
    [System.Diagnostics.CodeAnalysis.RequiresUnreferencedCode("Uses reflection over an arbitrary runtime type. For NativeAOT, call FrontMatter<T> or pass a dictionary.")]
    public MarkdownDocument FrontMatter(object data) {
        _frontMatter = FrontMatterBlock.FromObject(data);
        MarkdownObjectTreeBinder.BindDocument(this);
        return this;
    }

    /// <summary>Sets YAML front matter from a type known at publish time.</summary>
    public MarkdownDocument FrontMatter<
        [System.Diagnostics.CodeAnalysis.DynamicallyAccessedMembers(System.Diagnostics.CodeAnalysis.DynamicallyAccessedMemberTypes.PublicProperties)] T>(T data) {
        _frontMatter = FrontMatterBlock.FromObject(data);
        MarkdownObjectTreeBinder.BindDocument(this);
        return this;
    }

    /// <summary>Adds an MarkdownH1 heading.</summary>
    public MarkdownDocument H1(string text) => Append(new HeadingBlock(1, text));
    /// <summary>Adds an MarkdownH2 heading.</summary>
    public MarkdownDocument H2(string text) => Append(new HeadingBlock(2, text));
    /// <summary>Adds an MarkdownH3 heading.</summary>
    public MarkdownDocument H3(string text) => Append(new HeadingBlock(3, text));
    /// <summary>Adds an H4 heading.</summary>
    public MarkdownDocument H4(string text) => Append(new HeadingBlock(4, text));
    /// <summary>Adds an H5 heading.</summary>
    public MarkdownDocument H5(string text) => Append(new HeadingBlock(5, text));
    /// <summary>Adds an H6 heading.</summary>
    public MarkdownDocument H6(string text) => Append(new HeadingBlock(6, text));

    /// <summary>Adds a paragraph with plain text.</summary>
    public MarkdownDocument P(string text) => Append(new ParagraphBlock(new InlineSequence().Text(text)));

    /// <summary>Adds a paragraph composed with the paragraph builder.</summary>
    public MarkdownDocument P(Action<MarkdownParagraphBuilder> build) {
        MarkdownParagraphBuilder builder = new MarkdownParagraphBuilder();
        build(builder);
        return Append(new ParagraphBlock(builder.Inlines));
    }

    /// <summary>Adds a simple block quote with a single line of text.</summary>
    public MarkdownDocument Quote(string text) => Append(new QuoteBlock(new[] { text ?? string.Empty }));

    /// <summary>Adds a block quote composed via <see cref="MarkdownQuoteBuilder"/>.</summary>
    public MarkdownDocument Quote(Action<MarkdownQuoteBuilder> build) {
        MarkdownQuoteBuilder builder = new MarkdownQuoteBuilder();
        build(builder);
        return Append(builder.Build());
    }

    /// <summary>Adds a collapsible details block.</summary>
    public MarkdownDocument Details(string summary, Action<MarkdownDocument> buildBody, bool open = false) {
        var inner = MarkdownDocument.Create();
        buildBody(inner);
        var details = new DetailsBlock(new SummaryBlock(summary), inner.Blocks, open);
        return Append(details);
    }

    /// <summary>Adds a horizontal rule.</summary>
    public MarkdownDocument Hr() => Append(new HorizontalRuleBlock());

    /// <summary>Adds a callout/admonition block (Docs-style).</summary>
    public MarkdownDocument Callout(string kind, string title, string body) => Append(new CalloutBlock(kind, title, body));

    /// <summary>Adds a callout/admonition block with structured markdown body content.</summary>
    public MarkdownDocument Callout(string kind, string title, Action<MarkdownDocument> buildBody) {
        if (buildBody == null) {
            throw new ArgumentNullException(nameof(buildBody));
        }

        var body = MarkdownDocument.Create();
        buildBody(body);
        return Append(new CalloutBlock(kind, title, body.Blocks));
    }

    /// <summary>Adds an unordered list.</summary>
    public MarkdownDocument Ul(Action<MarkdownUnorderedListBuilder> build) {
        MarkdownUnorderedListBuilder builder = new MarkdownUnorderedListBuilder();
        build(builder);
        return Append(builder.Build());
    }

    /// <summary>Adds an ordered list.</summary>
    public MarkdownDocument Ol(Action<MarkdownOrderedListBuilder> build) {
        MarkdownOrderedListBuilder builder = new MarkdownOrderedListBuilder();
        build(builder);
        return Append(builder.Build());
    }

    /// <summary>Adds a definition list.</summary>
    public MarkdownDocument Dl(Action<MarkdownDefinitionListBuilder> build) {
        MarkdownDefinitionListBuilder builder = new MarkdownDefinitionListBuilder();
        build(builder);
        return Append(builder.Build());
    }

    /// <summary>Adds a fenced code block.</summary>
    public MarkdownDocument Code(string language, string content) {
        CodeBlock code = new CodeBlock(language, content);
        Append(code);
        return this;
    }

    /// <summary>Sets a caption for the last captionable block (image/code), or appends a paragraph.</summary>
    public MarkdownDocument Caption(string caption) {
        if (_lastBlock is ICaptionable cap) {
            cap.Caption = caption;
        } else {
            // Fallback: render as simple paragraph text
            P(caption);
        }
        return this;
    }

    /// <summary>Adds an image block with optional alt text and title.</summary>
    public MarkdownDocument Image(string path, string? alt, string? title)
        => Append(new ImageBlock(path, alt, title));

    /// <summary>Adds an image block with optional alt text, title, and size hints.</summary>
    public MarkdownDocument Image(string path, string? alt = null, string? title = null, double? width = null, double? height = null)
        => Append(new ImageBlock(path, alt, title, width, height));

    /// <summary>Adds a table built with <see cref="MarkdownTableBuilder"/>.</summary>
    public MarkdownDocument Table(Action<MarkdownTableBuilder> build) {
        MarkdownTableBuilder tb = new MarkdownTableBuilder();
        build(tb);
        return Append(tb.Build());
    }

    /// <summary>
    /// Builds a table via <see cref="MarkdownTableBuilder"/>, then applies auto-alignment heuristics if requested.
    /// </summary>
    public MarkdownDocument TableAuto(Action<MarkdownTableBuilder> build, bool alignNumeric = true, bool alignDates = true) {
        MarkdownTableBuilder tb = new MarkdownTableBuilder();
        build(tb);
        if (alignDates) tb.AlignDatesCenter();
        if (alignNumeric) tb.AlignNumericRight();
        return Append(tb.Build());
    }

    /// <summary>Adds an unordered list from a sequence of items using ToString().</summary>
    public MarkdownDocument Ul<T>(System.Collections.Generic.IEnumerable<T> items) {
        MarkdownUnorderedListBuilder builder = new MarkdownUnorderedListBuilder();
        builder.Items(items, null);
        return Append(builder.Build());
    }

    /// <summary>Adds an ordered list from a sequence of items using ToString().</summary>
    public MarkdownDocument Ol<T>(System.Collections.Generic.IEnumerable<T> items, int start = 1) {
        MarkdownOrderedListBuilder builder = new MarkdownOrderedListBuilder().StartAt(start);
        builder.Items(items, null);
        return Append(builder.Build());
    }

    /// <summary>Renders an embeddable HTML fragment. Wraps in &lt;article class="markdown-body"&gt; by default.</summary>
    internal string RenderHtmlFragment(MarkdownHtmlExportOptions? options = null) {
        MarkdownHtmlExportOptions operation = (options ?? new MarkdownHtmlExportOptions()).CloneForRender(MarkdownHtmlKind.Fragment);
        return HtmlRenderer.Render(this, operation);
    }

    /// <summary>
    /// Renders an embeddable HTML fragment and inlines the computed CSS and tiny scripts at the top/bottom
    /// of the fragment. Useful when you want a single self-contained chunk without a full HTML document.
    /// </summary>
    internal string RenderHtmlFragmentWithCss(MarkdownHtmlExportOptions? options = null) {
        MarkdownHtmlExportOptions operation = (options ?? new MarkdownHtmlExportOptions()).CloneForRender(MarkdownHtmlKind.Fragment);
        var parts = HtmlRenderer.RenderParts(this, operation);
        var sb = new StringBuilder();
        if (!string.IsNullOrEmpty(parts.Css)) sb.Append("<style>\n").Append(parts.Css).Append("\n</style>");
        sb.Append(parts.Body);
        if (!string.IsNullOrEmpty(parts.Scripts)) sb.Append("<script>\n").Append(parts.Scripts).Append("\n</script>");
        return sb.ToString();
    }

    /// <summary>Renders a standalone HTML5 document with optional CSS/JS assets.</summary>
    internal string RenderHtmlDocument(MarkdownHtmlExportOptions? options = null) {
        MarkdownHtmlExportOptions operation = (options ?? new MarkdownHtmlExportOptions()).CloneForRender(MarkdownHtmlKind.Document);
        return HtmlRenderer.Render(this, operation);
    }

    /// <summary>Returns rendered parts for advanced embedding (Head, Body, Css, Scripts).</summary>
    public MarkdownHtmlParts ToHtmlParts(MarkdownHtmlExportOptions? options = null) {
        MarkdownHtmlExportOptions operation = (options ?? new MarkdownHtmlExportOptions()).CloneForRender();
        return HtmlRenderer.RenderParts(this, operation);
    }

    internal (System.Collections.Generic.List<IMarkdownBlock> Blocks, MarkdownHeadingCatalog HeadingCatalog) GetBlocksAndHeadingSlugs(
        MarkdownHeadingIdentifierStyle headingIdentifierStyle = MarkdownHeadingIdentifierStyle.OfficeIMO,
        System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var registry = MarkdownSlug.CreateRegistry();
        var (realized, headingCatalog) = RealizeTocPlaceholders(registry, headingIdentifierStyle, cancellationToken);
        return (realized, headingCatalog);
    }





    /// <summary>
    /// Generates a MarkdownTable of Contents from headings already present in the document and inserts it.
    /// </summary>
    /// <param name="configure">Optional TOC options.</param>
    /// <param name="placeAtTop">When true, inserts TOC as the first block; otherwise appended.</param>
    public MarkdownDocument Toc(System.Action<MarkdownTocOptions>? configure = null, bool placeAtTop = false) {
        var opts = new MarkdownTocOptions();
        configure?.Invoke(opts);

        var placeholder = new TocPlaceholderBlock(opts, titleHeadingAlreadyInserted: opts.IncludeTitle);
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
        MarkdownObjectTreeBinder.BindDocument(this);
        return this;
    }

    /// <summary>
    /// Convenience helper to insert a MarkdownTable of Contents at the top with common parameters.
    /// </summary>
    public MarkdownDocument TocAtTop(string title = "Contents", int min = 1, int max = 3, bool ordered = false, int titleLevel = 2) {
        return Toc(opts => { opts.Title = title; opts.MinLevel = min; opts.MaxLevel = max; opts.Ordered = ordered; opts.TitleLevel = titleLevel; }, placeAtTop: true);
    }

    /// <summary>
    /// Inserts a TOC placeholder at the current position without a title heading by default.
    /// </summary>
    public MarkdownDocument TocHere(System.Action<MarkdownTocOptions>? configure = null) {
        var opts = new MarkdownTocOptions { IncludeTitle = false };
        configure?.Invoke(opts);
        var placeholder = new TocPlaceholderBlock(opts);
        _blocks.Add(placeholder);
        _lastBlock = placeholder;
        MarkdownObjectTreeBinder.BindDocument(this);
        return this;
    }

    /// <summary>
    /// Inserts a section TOC for the nearest preceding heading. Useful to place a small TOC under a section.
    /// </summary>
    public MarkdownDocument TocForPreviousHeading(string? title = "Contents", int min = 2, int max = 6, bool ordered = false, int titleLevel = 3) {
        return Toc(opts => { opts.Title = title ?? ""; opts.IncludeTitle = !string.IsNullOrEmpty(title); opts.MinLevel = min; opts.MaxLevel = max; opts.Ordered = ordered; opts.TitleLevel = titleLevel; opts.Scope = MarkdownTocScope.PreviousHeading; }, placeAtTop: false);
    }

    /// <summary>
    /// Inserts a section TOC scoped to the named heading.
    /// </summary>
    public MarkdownDocument TocForSection(string headingTitle, string? title = "Contents", int min = 2, int max = 6, bool ordered = false, int titleLevel = 3) {
        return Toc(opts => { opts.Title = title ?? ""; opts.IncludeTitle = !string.IsNullOrEmpty(title); opts.MinLevel = min; opts.MaxLevel = max; opts.Ordered = ordered; opts.TitleLevel = titleLevel; opts.Scope = MarkdownTocScope.HeadingTitle; opts.ScopeHeadingTitle = headingTitle; }, placeAtTop: false);
    }

    private (System.Collections.Generic.List<IMarkdownBlock> Blocks, MarkdownHeadingCatalog HeadingCatalog) RealizeTocPlaceholders(
        System.Collections.Generic.Dictionary<string, int> slugRegistry,
        MarkdownHeadingIdentifierStyle headingIdentifierStyle, System.Threading.CancellationToken cancellationToken) {
        // Create a shallow copy first
        var realized = new System.Collections.Generic.List<IMarkdownBlock>(_blocks);
        var headingCatalog = MarkdownHeadingCatalog.Create(realized, slugRegistry, headingIdentifierStyle, cancellationToken);
        // Replace placeholders with generated TOC blocks
        for (int i = 0; i < realized.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (realized[i] is ITocPlaceholderMarkdownBlock tocPlaceholder) {
                realized[i] = tocPlaceholder.RealizeToc(realized, i, headingCatalog);
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        return (realized, headingCatalog);
    }

    private static string? RenderFrontMatter(IFrontMatterMarkdownBlock? frontMatter, MarkdownSaveOptions options) {
        if (frontMatter == null) {
            return null;
        }

        return options.FrontMatterRendering switch {
            MarkdownFrontMatterRenderingMode.Preserve => frontMatter.RenderFrontMatter(),
            MarkdownFrontMatterRenderingMode.Omit => null,
            _ => throw new ArgumentOutOfRangeException(nameof(options.FrontMatterRendering), options.FrontMatterRendering, "Unknown front matter rendering mode.")
        };
    }

    private static bool AppendParseOwnedAbbreviationDefinitions(StringBuilder sb, MarkdownParseInfo? parseResult, MarkdownSaveOptions options) {
        if (options.AbbreviationDefinitionRendering == MarkdownAbbreviationDefinitionRenderingMode.Omit) {
            return false;
        }

        if (options.AbbreviationDefinitionRendering != MarkdownAbbreviationDefinitionRenderingMode.Preserve) {
            throw new ArgumentOutOfRangeException(nameof(options.AbbreviationDefinitionRendering), options.AbbreviationDefinitionRendering, "Unknown abbreviation definition rendering mode.");
        }

        if (parseResult == null || parseResult.AbbreviationDefinitions.Count == 0) {
            return false;
        }

        var wroteDefinition = false;
        for (int i = 0; i < parseResult.AbbreviationDefinitions.Count; i++) {
            var definition = parseResult.AbbreviationDefinitions[i];
            if (definition == null
                || string.IsNullOrWhiteSpace(definition.Label)
                || definition.IsListItemDefinition) {
                continue;
            }

            sb.AppendLine(RenderAbbreviationDefinition(parseResult, definition));
            wroteDefinition = true;
        }

        return wroteDefinition;
    }

    private static string RenderAbbreviationDefinition(MarkdownParseInfo parseResult, MarkdownAbbreviationDefinition definition) {
        if (definition.SourceSpan.HasValue
            && parseResult.TryCreateSourceSlice(definition.SourceSpan.Value, out var slice)) {
            var sourceLine = slice.Text.TrimEnd('\r', '\n');
            if (!string.IsNullOrWhiteSpace(sourceLine)) {
                return sourceLine;
            }
        }

        return "*[" + definition.Label + "]: " + definition.Title;
    }

    private static string NormalizeLineEndings(string value, string lineEnding) {
        if (string.IsNullOrEmpty(value)) {
            return string.Empty;
        }

        var normalized = value.Replace("\r\n", "\n").Replace('\r', '\n');
        return lineEnding == "\n"
            ? normalized
            : normalized.Replace("\n", lineEnding);
    }
}
