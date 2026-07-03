namespace OfficeIMO.Markdown;

/// <summary>
/// Sequence of inline nodes used in paragraphs and list items.
/// </summary>
public sealed class InlineSequence : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline {
    private readonly List<IMarkdownInline> _inlines = new List<IMarkdownInline>();
    private readonly IReadOnlyList<object> _itemsView;

    /// <summary>Creates an empty inline sequence.</summary>
    public InlineSequence() {
        _itemsView = new InlineObjectReadOnlyList(_inlines);
    }

    // When composing via the fluent/builder APIs, auto-spacing between adjacent inline nodes is convenient.
    // When parsing Markdown source, spacing is already present in TextRun nodes, so auto-spacing would double spaces.
    internal bool AutoSpacing { get; set; } = true;
    /// <summary>Exposes the inline nodes for safe iteration.</summary>
    public IReadOnlyList<IMarkdownInline> Nodes => _inlines;
    /// <summary>Legacy object-typed inline view retained for compatibility.</summary>
    public IReadOnlyList<object> Items => _itemsView;
    /// <summary>Adds plain text.</summary>
    public InlineSequence Text(string text) { _inlines.Add(new TextRun(text)); return this; }
    /// <summary>Adds a hyperlink.</summary>
    public InlineSequence Link(string text, string url, string? title = null, string? linkTarget = null, string? linkRel = null) { _inlines.Add(new LinkInline(text, url, title, linkTarget, linkRel)); return this; }
    /// <summary>Adds bold text.</summary>
    public InlineSequence Bold(string text) { _inlines.Add(new BoldInline(text)); return this; }
    /// <summary>Adds bold+italic text.</summary>
    public InlineSequence BoldItalic(string text) { _inlines.Add(new BoldItalicInline(text)); return this; }
    /// <summary>Adds italic text.</summary>
    public InlineSequence Italic(string text) { _inlines.Add(new ItalicInline(text)); return this; }
    /// <summary>Adds inline code.</summary>
    public InlineSequence Code(string text) { _inlines.Add(new CodeSpanInline(text)); return this; }
    /// <summary>Adds a footnote reference (e.g., [^id]).</summary>
    public InlineSequence FootnoteRef(string label) { _inlines.Add(new FootnoteRefInline(label)); return this; }
    /// <summary>Adds strikethrough text.</summary>
    public InlineSequence Strike(string text) { _inlines.Add(new StrikethroughInline(text)); return this; }
    /// <summary>Adds highlighted text rendered as <c>==text==</c>.</summary>
    public InlineSequence Highlight(string text) { _inlines.Add(new HighlightInline(text)); return this; }
    /// <summary>Adds underlined text (HTML-only in Markdown).</summary>
    public InlineSequence Underline(string text) { _inlines.Add(new UnderlineInline(text)); return this; }
    /// <summary>Adds superscript text rendered as <c>^text^</c> in Markdown and <c>&lt;sup&gt;</c> in HTML.</summary>
    public InlineSequence Superscript(string text) { _inlines.Add(new SuperscriptInline(text)); return this; }
    /// <summary>Adds subscript text rendered as <c>~text~</c> in Markdown and <c>&lt;sub&gt;</c> in HTML.</summary>
    public InlineSequence Subscript(string text) { _inlines.Add(new SubscriptInline(text)); return this; }
    /// <summary>Adds inserted text rendered as <c>++text++</c> in Markdown and <c>&lt;ins&gt;</c> in HTML.</summary>
    public InlineSequence Inserted(string text) { _inlines.Add(new InsertedInline(text)); return this; }
    /// <summary>Adds quoted text rendered via inline HTML.</summary>
    public InlineSequence Quote(string text) { _inlines.Add(new HtmlTagSequenceInline("q", new InlineSequence().Text(text))); return this; }
    /// <summary>Adds a linked image (useful for badges).</summary>
    public InlineSequence ImageLink(string alt, string imageUrl, string linkUrl, string? title = null, string? linkTitle = null) { _inlines.Add(new ImageLinkInline(alt, imageUrl, linkUrl, title, linkTitle)); return this; }
    /// <summary>Adds a standalone inline image.</summary>
    public InlineSequence Image(string alt, string src, string? title = null) { _inlines.Add(new ImageInline(alt, src, title)); return this; }
    /// <summary>Adds a hard line break.</summary>
    public InlineSequence HardBreak() { _inlines.Add(new HardBreakInline()); return this; }
    /// <summary>Adds a soft line break.</summary>
    public InlineSequence SoftBreak() { _inlines.Add(new SoftBreakInline()); return this; }

    // Internal escape hatch for the reader to attach richer inline nodes without expanding the public fluent API.
    internal InlineSequence AddRaw(IMarkdownInline node) { if (node != null) _inlines.Add(node); return this; }

    /// <summary>
    /// Replaces the inline nodes in this sequence.
    /// Extension authors can use this from reader transform hooks to normalize a parsed inline AST
    /// while preserving source spans on any existing node instances they keep.
    /// </summary>
    public void ReplaceItems(IEnumerable<IMarkdownInline> nodes) {
        if (nodes == null) {
            _inlines.Clear();
            return;
        }

        var replacement = nodes.Where(node => node != null).ToArray();
        _inlines.Clear();
        for (int i = 0; i < replacement.Length; i++) {
            _inlines.Add(replacement[i]);
        }
    }

    internal string RenderMarkdown() {
        StringBuilder sb = new StringBuilder();
        var options = MarkdownRenderContext.Options;
        MarkdownInlineMarkdownRenderContext? context = options == null
            ? null
            : new MarkdownInlineMarkdownRenderContext(options, MarkdownRenderContext.WriteContext);
        for (int i = 0; i < _inlines.Count; i++) {
            if (AutoSpacing && i > 0) {
                var prev = _inlines[i - 1];
                var cur = _inlines[i];
                if (prev is not HardBreakInline && cur is not HardBreakInline &&
                    prev is not SoftBreakInline && cur is not SoftBreakInline) sb.Append(' ');
            }
            sb.Append(RenderMarkdown(_inlines[i], context));
        }
        return sb.ToString();
    }

    internal string RenderMarkdownWithTextEscaper(Func<string?, string> textEscaper) {
        if (textEscaper == null) {
            return RenderMarkdown();
        }

        StringBuilder sb = new StringBuilder();
        var options = MarkdownRenderContext.Options;
        MarkdownInlineMarkdownRenderContext? context = options == null
            ? null
            : new MarkdownInlineMarkdownRenderContext(options, MarkdownRenderContext.WriteContext);
        for (int i = 0; i < _inlines.Count; i++) {
            if (AutoSpacing && i > 0) {
                var prev = _inlines[i - 1];
                var cur = _inlines[i];
                if (prev is not HardBreakInline && cur is not HardBreakInline &&
                    prev is not SoftBreakInline && cur is not SoftBreakInline) sb.Append(' ');
            }

            sb.Append(RenderMarkdown(_inlines[i], context, textEscaper));
        }

        return sb.ToString();
    }

    internal string RenderHtml() {
        StringBuilder sb = new StringBuilder();
        var options = HtmlRenderContext.Options;
        MarkdownInlineHtmlRenderContext? context = options == null
            ? null
            : new MarkdownInlineHtmlRenderContext(options, HtmlRenderContext.BodyContext);
        for (int i = 0; i < _inlines.Count; i++) {
            if (AutoSpacing && i > 0) {
                var prev = _inlines[i - 1];
                var cur = _inlines[i];
                if (prev is not HardBreakInline && cur is not HardBreakInline &&
                    prev is not SoftBreakInline && cur is not SoftBreakInline) sb.Append(' ');
            }
            sb.Append(RenderHtml(_inlines[i], options, context));
        }
        return sb.ToString();
    }

    string IRenderableMarkdownInline.RenderMarkdown() => RenderMarkdown();
    string IRenderableMarkdownInline.RenderHtml() => RenderHtml();
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
            TextRun text => textEscaper(text.Text),
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
            var extension = extensions[i];
            if (extension == null || !extension.Matches(syntaxNode)) {
                continue;
            }

            var rendered = extension.RenderMarkdown(node, syntaxNode, context);
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
            var extension = extensions[i];
            if (extension == null || !extension.Matches(node)) {
                continue;
            }

            var rendered = extension.RenderMarkdownWithContext(node, context);
            if (rendered != null) {
                return rendered;
            }
        }

        return null;
    }

    private static string RenderHtml(IMarkdownInline node, HtmlOptions? options, MarkdownInlineHtmlRenderContext? context) {
        var overridden = TryRenderInlineSyntaxOverride(node, context);
        if (overridden != null) {
            return overridden;
        }

        overridden = TryRenderInlineOverride(node, context);
        if (overridden != null) {
            return overridden;
        }

        if (options != null && node is IContextualHtmlMarkdownInline contextualInline) {
            return MarkdownInlineAttributeRenderer.RenderHtml(node, contextualInline.RenderHtml(options), options);
        }

        return MarkdownInlineAttributeRenderer.RenderHtml(node, GetRenderable(node).RenderHtml(), options);
    }

    private static string? TryRenderInlineSyntaxOverride(IMarkdownInline node, MarkdownInlineHtmlRenderContext? context) {
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

        var extensions = context.Options.InlineRenderExtensions;
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

internal sealed class InlineObjectReadOnlyList : IReadOnlyList<object> {
    private readonly IReadOnlyList<IMarkdownInline> _nodes;

    public InlineObjectReadOnlyList(IReadOnlyList<IMarkdownInline> nodes) {
        _nodes = nodes ?? throw new ArgumentNullException(nameof(nodes));
    }

    public int Count => _nodes.Count;

    public object this[int index] => _nodes[index];

    public IEnumerator<object> GetEnumerator() {
        for (int i = 0; i < _nodes.Count; i++) {
            yield return _nodes[i];
        }
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
}
