using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Markdown;

/// <summary>
/// Root document container and fluent API entrypoint for composing Markdown.
/// Supports a fluent-chaining style and an explicit object model via <see cref="Add(IMarkdownBlock)"/>.
/// </summary>
public class MarkdownDoc {
    private readonly List<IMarkdownBlock> _blocks = new();
    private IMarkdownBlock? _lastBlock;
    private FrontMatterBlock? _frontMatter;

    /// <summary>Creates a new, empty Markdown document.</summary>
    public static MarkdownDoc Create() => new MarkdownDoc();

    /// <summary>All blocks added to the document (excluding front matter).</summary>
    public IReadOnlyList<IMarkdownBlock> Blocks => _blocks;

    /// <summary>Adds a block instance (object-model style).</summary>
    /// <param name="block">Block to append to the document.</param>
    /// <returns>Same <see cref="MarkdownDoc"/> for chaining.</returns>
    public MarkdownDoc Add(IMarkdownBlock block) {
        if (block is FrontMatterBlock fm) {
            _frontMatter = fm;
        } else {
            _blocks.Add(block);
            _lastBlock = block;
        }
        return this;
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
    public MarkdownDoc Image(string path, string? alt = null, string? title = null) => Add(new ImageBlock(path, alt, title));

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
        var blocks = RealizeTocPlaceholders();
        StringBuilder sb = new StringBuilder();
        if (_frontMatter != null) {
            sb.AppendLine(_frontMatter.Render());
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
        return Utilities.HtmlRenderer.Render(this, options);
    }

    /// <summary>Renders a standalone HTML5 document with optional CSS/JS assets.</summary>
    public string ToHtmlDocument(HtmlOptions? options = null) {
        options ??= new HtmlOptions { Kind = HtmlKind.Document };
        options.Kind = HtmlKind.Document;
        return Utilities.HtmlRenderer.Render(this, options);
    }

    /// <summary>Asynchronously renders an embeddable HTML fragment.</summary>
    public System.Threading.Tasks.Task<string> ToHtmlFragmentAsync(HtmlOptions? options = null) => System.Threading.Tasks.Task.FromResult(ToHtmlFragment(options));
    /// <summary>Asynchronously renders a full HTML document.</summary>
    public System.Threading.Tasks.Task<string> ToHtmlDocumentAsync(HtmlOptions? options = null) => System.Threading.Tasks.Task.FromResult(ToHtmlDocument(options));

    /// <summary>Returns rendered parts for advanced embedding (Head, Body, Css, Scripts).</summary>
    public HtmlRenderParts ToHtmlParts(HtmlOptions? options = null) {
        options ??= new HtmlOptions { Kind = HtmlKind.Fragment };
        return Utilities.HtmlRenderer.RenderParts(this, options);
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
        await Utilities.FileCompat.WriteAllTextAsync(path, html, System.Text.Encoding.UTF8).ConfigureAwait(false);
        if (!string.IsNullOrEmpty(options.ExternalCssOutputPath) && options._externalCssContentToWrite is not null) {
            await Utilities.FileCompat.WriteAllTextAsync(options.ExternalCssOutputPath!, options._externalCssContentToWrite, System.Text.Encoding.UTF8).ConfigureAwait(false);
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

    private System.Collections.Generic.List<IMarkdownBlock> RealizeTocPlaceholders() {
        // Create a shallow copy first
        var realized = new System.Collections.Generic.List<IMarkdownBlock>(_blocks);
        // Collect heading info from realized list with indices
        var headings = new System.Collections.Generic.List<(int Index, int Level, string Text)>();
        for (int idx = 0; idx < realized.Count; idx++) {
            if (realized[idx] is HeadingBlock h) headings.Add((idx, h.Level, h.Text));
        }
        // Replace placeholders with generated TOC blocks
        for (int i = 0; i < realized.Count; i++) {
            if (realized[i] is TocPlaceholderBlock tp) {
                var opts = tp.Options;
                var toc = new TocBlock { Ordered = opts.Ordered };
                // Determine scope bounds
                int startIdx = 0; int endIdx = realized.Count;
                if (opts.Scope == TocScope.PreviousHeading) {
                    // Root at the nearest preceding heading with level < MinLevel if available; otherwise nearest heading.
                    var prev = headings.LastOrDefault(h => h.Index < i && h.Level < opts.MinLevel);
                    if (prev == default) prev = headings.LastOrDefault(h => h.Index < i);
                    if (prev != default) {
                        startIdx = prev.Index + 1;
                        var nextAtOrAbove = headings.FirstOrDefault(h => h.Index > prev.Index && h.Level <= prev.Level);
                        if (nextAtOrAbove != default) endIdx = nextAtOrAbove.Index;
                    }
                } else if (opts.Scope == TocScope.HeadingTitle && !string.IsNullOrWhiteSpace(opts.ScopeHeadingTitle)) {
                    var root = headings.FirstOrDefault(h => string.Equals(h.Text, opts.ScopeHeadingTitle, System.StringComparison.OrdinalIgnoreCase));
                    if (root != default) {
                        startIdx = root.Index + 1;
                        var nextAtOrAbove = headings.FirstOrDefault(h => h.Index > root.Index && h.Level <= root.Level);
                        if (nextAtOrAbove != default) endIdx = nextAtOrAbove.Index;
                    }
                }
                foreach (var h in headings) {
                    if (h.Index < startIdx || h.Index >= endIdx) continue;
                    if (h.Level < opts.MinLevel || h.Level > opts.MaxLevel) continue;
                    var anchor = MarkdownSlug.GitHub(h.Text);
                    toc.Entries.Add(new TocBlock.Entry { Level = h.Level, Text = h.Text, Anchor = anchor });
                }
                realized[i] = toc;
            }
        }
        return realized;
    }
}
