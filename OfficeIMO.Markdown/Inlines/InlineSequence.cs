namespace OfficeIMO.Markdown;

/// <summary>
/// Sequence of inline nodes used in paragraphs and list items.
/// </summary>
public sealed class InlineSequence : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, IHtmlAppendableMarkdownInline, IMarkdownAppendableMarkdownInline {
    private readonly InlineNodeList _inlines;

    /// <summary>Creates an empty inline sequence.</summary>
    public InlineSequence() {
        _inlines = new InlineNodeList();
    }

    internal InlineSequence(int capacity) {
        _inlines = new InlineNodeList(Math.Max(0, capacity));
    }

    // When composing via the fluent/builder APIs, auto-spacing between adjacent inline nodes is convenient.
    // When parsing Markdown source, spacing is already present in MarkdownTextRun nodes, so auto-spacing would double spaces.
    internal bool AutoSpacing { get; set; } = true;
    /// <summary>Exposes the inline nodes for safe iteration.</summary>
    public IReadOnlyList<IMarkdownInline> Nodes => _inlines;
    /// <summary>Adds plain text.</summary>
    public InlineSequence Text(string text) { AppendNode(new MarkdownTextRun(text)); return this; }
    /// <summary>Adds a hyperlink.</summary>
    public InlineSequence Link(string text, string url, string? title = null, string? linkTarget = null, string? linkRel = null) { AppendNode(new LinkInline(text, url, title, linkTarget, linkRel)); return this; }
    /// <summary>Adds bold text.</summary>
    public InlineSequence Bold(string text) { AppendNode(new BoldInline(text)); return this; }
    /// <summary>Adds bold+italic text.</summary>
    public InlineSequence BoldItalic(string text) { AppendNode(new BoldItalicInline(text)); return this; }
    /// <summary>Adds italic text.</summary>
    public InlineSequence Italic(string text) { AppendNode(new ItalicInline(text)); return this; }
    /// <summary>Adds inline code.</summary>
    public InlineSequence Code(string text) { AppendNode(new CodeSpanInline(text)); return this; }
    /// <summary>Adds a footnote reference (e.g., [^id]).</summary>
    public InlineSequence FootnoteRef(string label) { AppendNode(new FootnoteRefInline(label)); return this; }
    /// <summary>Adds strikethrough text.</summary>
    public InlineSequence Strikethrough(string text) { AppendNode(new StrikethroughInline(text)); return this; }
    /// <summary>Adds highlighted text rendered as <c>==text==</c>.</summary>
    public InlineSequence Highlight(string text) { AppendNode(new HighlightInline(text)); return this; }
    /// <summary>Adds underlined text (HTML-only in Markdown).</summary>
    public InlineSequence Underline(string text) { AppendNode(new UnderlineInline(text)); return this; }
    /// <summary>Adds superscript text rendered as <c>^text^</c> in Markdown and <c>&lt;sup&gt;</c> in HTML.</summary>
    public InlineSequence Superscript(string text) { AppendNode(new SuperscriptInline(text)); return this; }
    /// <summary>Adds subscript text rendered as <c>~text~</c> in Markdown and <c>&lt;sub&gt;</c> in HTML.</summary>
    public InlineSequence Subscript(string text) { AppendNode(new SubscriptInline(text)); return this; }
    /// <summary>Adds inserted text rendered as <c>++text++</c> in Markdown and <c>&lt;ins&gt;</c> in HTML.</summary>
    public InlineSequence Inserted(string text) { AppendNode(new InsertedInline(text)); return this; }
    /// <summary>Adds quoted text rendered via inline HTML.</summary>
    public InlineSequence Quote(string text) { AppendNode(new HtmlTagSequenceInline("q", new InlineSequence().Text(text))); return this; }
    /// <summary>Adds a linked image (useful for badges).</summary>
    public InlineSequence ImageLink(string alt, string imageUrl, string linkUrl, string? title = null, string? linkTitle = null) { AppendNode(new ImageLinkInline(alt, imageUrl, linkUrl, title, linkTitle)); return this; }
    /// <summary>Adds a standalone inline image.</summary>
    public InlineSequence Image(string alt, string src, string? title = null) { AppendNode(new ImageInline(alt, src, title)); return this; }
    /// <summary>Adds a hard line break.</summary>
    public InlineSequence HardBreak() { AppendNode(new HardBreakInline()); return this; }
    /// <summary>Adds a soft line break.</summary>
    public InlineSequence SoftBreak() { AppendNode(new SoftBreakInline()); return this; }

    // Internal escape hatch for the reader to attach richer inline nodes without expanding the public fluent API.
    internal InlineSequence AddRaw(IMarkdownInline node) { if (node != null) AppendNode(node); return this; }

    /// <summary>
    /// Replaces the inline nodes in this sequence.
    /// Extension authors can use this from reader transform hooks to normalize a parsed inline AST
    /// while preserving source spans on any existing node instances they keep.
    /// </summary>
    public void ReplaceItems(IEnumerable<IMarkdownInline> nodes) {
        if (nodes == null) {
            _inlines.Reset();
            return;
        }

        var replacement = nodes.Where(node => node != null).ToArray();
        _inlines.Reset();
        for (int i = 0; i < replacement.Length; i++) {
            _inlines.Append(replacement[i]);
        }

        if (Parent != null) {
            MarkdownObjectTreeBinder.RebindSubtree(this);
        }
    }

    // A sequence that is already part of a bound tree places each new node as it is added, so editing a
    // block after it was added to a document does not depend on a later full re-bind.
    private void AppendNode(IMarkdownInline node) {
        _inlines.Append(node);
        if (Parent != null) {
            MarkdownObjectTreeBinder.BindAppendedInline(this, node);
        }
    }

    internal string RenderMarkdown() {
        StringBuilder sb = MarkdownStringBuilderCache.Acquire();
        AppendMarkdown(sb);
        return MarkdownStringBuilderCache.GetStringAndRelease(sb);
    }

    internal void AppendMarkdown(StringBuilder sb) {
        MarkdownRenderContext.ThrowIfCancellationRequested();
        // Rendering descends nested sequences recursively; a hand-built tree deeper than the stack
        // fails with a catchable exception instead of ending the process.
        System.Runtime.CompilerServices.RuntimeHelpers.EnsureSufficientExecutionStack();
        MarkdownInlineMarkdownRenderContext? context = CreateMarkdownRenderContext();
        for (int i = 0; i < _inlines.Count; i++) {
            MarkdownRenderContext.ThrowIfCancellationRequested();
            if (AutoSpacing && i > 0) {
                var prev = _inlines[i - 1];
                var cur = _inlines[i];
                if (prev is not HardBreakInline && cur is not HardBreakInline &&
                    prev is not SoftBreakInline && cur is not SoftBreakInline) sb.Append(' ');
            }
            AppendMarkdown(sb, _inlines[i], context);
            MarkdownRenderContext.ThrowIfCancellationRequested();
        }
    }

    // Only render extensions read the inline context, so it is created when one is registered.
    private static MarkdownInlineMarkdownRenderContext? CreateMarkdownRenderContext() {
        var options = MarkdownRenderContext.Options;
        return options == null
            || ((options.InlineRenderExtensions?.Count ?? 0) == 0 && (options.SyntaxInlineRenderExtensions?.Count ?? 0) == 0)
            ? null
            : new MarkdownInlineMarkdownRenderContext(options, MarkdownRenderContext.WriteContext);
    }

    private static void AppendMarkdown(StringBuilder sb, IMarkdownInline node, MarkdownInlineMarkdownRenderContext? context) {
        var overridden = TryRenderInlineSyntaxMarkdownOverride(node, context)
            ?? TryRenderInlineMarkdownOverride(node, context);
        if (overridden != null) {
            sb.Append(overridden);
            return;
        }

        if (node is IMarkdownAppendableMarkdownInline appendable) {
            appendable.AppendMarkdown(sb);
            // Generic attributes follow the inline, as MarkdownInlineAttributeRenderer writes them.
            if (node is MarkdownObject markdownObject && !markdownObject.Attributes.IsEmpty) {
                sb.Append(MarkdownAttributeBlockRenderer.RenderInlineTrailing(markdownObject.Attributes));
            }

            return;
        }

        sb.Append(MarkdownInlineAttributeRenderer.RenderMarkdown(node, GetRenderable(node).RenderMarkdown()));
    }

    internal string RenderMarkdownWithTextEscaper(Func<string?, string> textEscaper) {
        MarkdownRenderContext.ThrowIfCancellationRequested();
        if (textEscaper == null) {
            return RenderMarkdown();
        }

        System.Runtime.CompilerServices.RuntimeHelpers.EnsureSufficientExecutionStack();
        StringBuilder sb = new StringBuilder();
        MarkdownInlineMarkdownRenderContext? context = CreateMarkdownRenderContext();
        for (int i = 0; i < _inlines.Count; i++) {
            MarkdownRenderContext.ThrowIfCancellationRequested();
            if (AutoSpacing && i > 0) {
                var prev = _inlines[i - 1];
                var cur = _inlines[i];
                if (prev is not HardBreakInline && cur is not HardBreakInline &&
                    prev is not SoftBreakInline && cur is not SoftBreakInline) sb.Append(' ');
            }

            sb.Append(RenderMarkdown(_inlines[i], context, textEscaper));
            MarkdownRenderContext.ThrowIfCancellationRequested();
        }

        return sb.ToString();
    }

    internal string RenderHtml() {
        if (_inlines.Count == 0) {
            return string.Empty;
        }

        var sb = new StringBuilder();
        AppendHtml(sb);
        return sb.ToString();
    }

    internal void AppendHtml(StringBuilder sb) {
        System.Runtime.CompilerServices.RuntimeHelpers.EnsureSufficientExecutionStack();
        var options = HtmlRenderContext.Options;
        // Only render extensions read the inline context, so it is created when one is registered.
        MarkdownInlineHtmlRenderContext? context = options == null
            || (options.InlineRenderExtensionsIfAny == null && options.SyntaxInlineRenderExtensionsIfAny == null)
            ? null
            : new MarkdownInlineHtmlRenderContext(options, HtmlRenderContext.BodyContext);
        for (int i = 0; i < _inlines.Count; i++) {
            if (AutoSpacing && i > 0) {
                var prev = _inlines[i - 1];
                var cur = _inlines[i];
                if (prev is not HardBreakInline && cur is not HardBreakInline &&
                    prev is not SoftBreakInline && cur is not SoftBreakInline) sb.Append(' ');
            }
            AppendHtml(sb, _inlines[i], options, context);
        }
    }

    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
    void IHtmlAppendableMarkdownInline.AppendHtml(StringBuilder builder) => AppendHtml(builder);
    void IMarkdownAppendableMarkdownInline.AppendMarkdown(StringBuilder builder) => AppendMarkdown(builder);
    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => InlinePlainText.AppendPlainText(sb, this);

    private static IRenderableMarkdownInline GetRenderable(IMarkdownInline node) {
        return node as IRenderableMarkdownInline
            ?? throw new InvalidOperationException($"Inline node of type '{node.GetType().FullName}' does not implement {nameof(IRenderableMarkdownInline)}.");
    }

    private static string RenderMarkdown(IMarkdownInline node, MarkdownInlineMarkdownRenderContext? context, Func<string?, string>? textEscaper = null) {
        var overridden = TryRenderInlineSyntaxMarkdownOverride(node, context);
        if (overridden != null) {
            return overridden;
        }

        overridden = TryRenderInlineMarkdownOverride(node, context);
        if (overridden != null) {
            return overridden;
        }

        if (textEscaper != null) {
            var rendered = RenderMarkdownWithEscapedTextRuns(node, textEscaper);
            if (rendered != null) {
                return MarkdownInlineAttributeRenderer.RenderMarkdown(node, rendered);
            }
        }

        return MarkdownInlineAttributeRenderer.RenderMarkdown(node, GetRenderable(node).RenderMarkdown());
    }

    private static string? RenderMarkdownWithEscapedTextRuns(IMarkdownInline node, Func<string?, string> textEscaper) {
        return node switch {
            MarkdownTextRun text => textEscaper(text.Text),
            BoldSequenceInline bold => "**" + bold.Inlines.RenderMarkdownWithTextEscaper(textEscaper) + "**",
            ItalicSequenceInline italic => "_" + italic.Inlines.RenderMarkdownWithTextEscaper(textEscaper) + "_",
            BoldItalicSequenceInline boldItalic => "***" + boldItalic.Inlines.RenderMarkdownWithTextEscaper(textEscaper) + "***",
            StrikethroughSequenceInline strike => "~~" + strike.Inlines.RenderMarkdownWithTextEscaper(textEscaper) + "~~",
            HighlightSequenceInline highlight => "==" + highlight.Inlines.RenderMarkdownWithTextEscaper(textEscaper) + "==",
            InsertedSequenceInline inserted => "++" + inserted.Inlines.RenderMarkdownWithTextEscaper(textEscaper) + "++",
            SuperscriptSequenceInline superscript => "^" + superscript.Inlines.RenderMarkdownWithTextEscaper(textEscaper) + "^",
            SubscriptSequenceInline subscript => "~" + subscript.Inlines.RenderMarkdownWithTextEscaper(textEscaper) + "~",
            HtmlTagSequenceInline htmlTag => "<" + htmlTag.TagName + ">" + htmlTag.Inlines.RenderMarkdownWithTextEscaper(textEscaper) + "</" + htmlTag.TagName + ">",
            LinkInline link => RenderLinkMarkdownWithEscapedText(link, textEscaper),
            ImageInline image => RenderImageMarkdownWithEscapedText(image, textEscaper),
            ImageLinkInline imageLink => RenderImageLinkMarkdownWithEscapedText(imageLink, textEscaper),
            _ => null
        };
    }

    private static string RenderLinkMarkdownWithEscapedText(LinkInline link, Func<string?, string> textEscaper) {
        var title = MarkdownEscaper.FormatOptionalTitle(link.Title);
        var label = link.LabelInlines != null
            ? link.LabelInlines.RenderMarkdownWithTextEscaper(textEscaper)
            : textEscaper(link.Text);
        return $"[{label}]({MarkdownEscaper.EscapeLinkUrl(link.Url)}{title})";
    }

    private static string RenderImageMarkdownWithEscapedText(ImageInline image, Func<string?, string> textEscaper) {
        if ((MarkdownRenderContext.Options?.ImageRenderingMode ?? MarkdownImageRenderingMode.RichMarkdown) == MarkdownImageRenderingMode.Html) {
            return image.RenderHtml();
        }

        var title = MarkdownEscaper.FormatOptionalTitle(image.Title);
        return $"![{textEscaper(image.PlainAlt)}]({MarkdownEscaper.EscapeImageSrc(image.Src)}{title})";
    }

    private static string RenderImageLinkMarkdownWithEscapedText(ImageLinkInline imageLink, Func<string?, string> textEscaper) {
        if ((MarkdownRenderContext.Options?.ImageRenderingMode ?? MarkdownImageRenderingMode.RichMarkdown) == MarkdownImageRenderingMode.Html) {
            return imageLink.RenderHtml();
        }

        var title = MarkdownEscaper.FormatOptionalTitle(imageLink.Title);
        var linkTitle = MarkdownEscaper.FormatOptionalTitle(imageLink.LinkTitle);
        return $"[![{textEscaper(imageLink.PlainAlt)}]({MarkdownEscaper.EscapeImageSrc(imageLink.ImageUrl)}{title})]({MarkdownEscaper.EscapeLinkUrl(imageLink.LinkUrl)}{linkTitle})";
    }

    private static string? TryRenderInlineSyntaxMarkdownOverride(IMarkdownInline node, MarkdownInlineMarkdownRenderContext? context) {
        if (context == null) {
            return null;
        }

        var extensions = context.Options.SyntaxInlineRenderExtensions;
        if (extensions == null || extensions.Count == 0) {
            return null;
        }

        var syntaxNode = context.FindSyntaxNode(node);
        if (syntaxNode == null) {
            return null;
        }

        for (int i = extensions.Count - 1; i >= 0; i--) {
            context.CancellationToken.ThrowIfCancellationRequested();
            var extension = extensions[i];
            if (extension == null || !extension.Matches(syntaxNode)) {
                continue;
            }

            var rendered = extension.RenderMarkdown(node, syntaxNode, context);
            context.CancellationToken.ThrowIfCancellationRequested();
            if (rendered != null) {
                return rendered;
            }
        }

        return null;
    }

    private static string? TryRenderInlineMarkdownOverride(IMarkdownInline node, MarkdownInlineMarkdownRenderContext? context) {
        if (context == null) {
            return null;
        }

        var extensions = context.Options.InlineRenderExtensions;
        if (extensions == null || extensions.Count == 0) {
            return null;
        }

        for (int i = extensions.Count - 1; i >= 0; i--) {
            context.CancellationToken.ThrowIfCancellationRequested();
            var extension = extensions[i];
            if (extension == null || !extension.Matches(node)) {
                continue;
            }

            var rendered = extension.RenderMarkdownWithContext(node, context);
            context.CancellationToken.ThrowIfCancellationRequested();
            if (rendered != null) {
                return rendered;
            }
        }

        return null;
    }

    private static void AppendHtml(StringBuilder sb, IMarkdownInline node, MarkdownHtmlExportOptions? options, MarkdownInlineHtmlRenderContext? context) {
        var overridden = TryRenderInlineSyntaxOverride(node, context);
        if (overridden != null) {
            sb.Append(overridden);
            return;
        }

        overridden = TryRenderInlineOverride(node, context);
        if (overridden != null) {
            sb.Append(overridden);
            return;
        }

        if (options != null && node is IContextualHtmlMarkdownInline contextualInline) {
            sb.Append(MarkdownInlineAttributeRenderer.RenderHtml(node, contextualInline.RenderHtml(options), options));
            return;
        }

        // Generic attributes are spliced into the rendered opening tag, so only attribute-free
        // nodes can write straight into the shared buffer.
        if (node is IHtmlAppendableMarkdownInline appendable
            && (node is not MarkdownObject markdownObject || markdownObject.Attributes.IsEmpty)) {
            appendable.AppendHtml(sb);
            return;
        }

        sb.Append(MarkdownInlineAttributeRenderer.RenderHtml(node, GetRenderable(node).RenderHtml(), options));
    }

    private static string? TryRenderInlineSyntaxOverride(IMarkdownInline node, MarkdownInlineHtmlRenderContext? context) {
        if (context == null) {
            return null;
        }

        var extensions = context.Options.SyntaxInlineRenderExtensionsIfAny;
        if (extensions == null || extensions.Count == 0) {
            return null;
        }

        var syntaxNode = context.FindSyntaxNode(node);
        if (syntaxNode == null) {
            return null;
        }

        for (int i = extensions.Count - 1; i >= 0; i--) {
            var extension = extensions[i];
            if (extension == null || !extension.Matches(syntaxNode)) {
                continue;
            }

            var rendered = extension.RenderHtml(node, syntaxNode, context);
            if (rendered != null) {
                return rendered;
            }
        }

        return null;
    }

    private static string? TryRenderInlineOverride(IMarkdownInline node, MarkdownInlineHtmlRenderContext? context) {
        if (context == null) {
            return null;
        }

        var extensions = context.Options.InlineRenderExtensionsIfAny;
        if (extensions == null || extensions.Count == 0) {
            return null;
        }

        for (int i = extensions.Count - 1; i >= 0; i--) {
            var extension = extensions[i];
            if (extension == null || !extension.Matches(node)) {
                continue;
            }

            var rendered = extension.RenderHtmlWithContext(node, context);
            if (rendered != null) {
                return rendered;
            }
        }

        return null;
    }
}
