using System.Runtime.CompilerServices;

namespace OfficeIMO.Markdown;

/// <summary>
/// MarkdownParagraph block containing a sequence of inline nodes.
/// </summary>
public sealed class ParagraphBlock : MarkdownBlock, IMarkdownBlock, IParagraphMarkdownBlock, IInlineSyntaxMarkdownBlock, ISyntaxMarkdownBlock, IHtmlAppendableMarkdownBlock {
    /// <summary>Inline content within this paragraph.</summary>
    public InlineSequence Inlines { get; }
    // Only paragraphs with trailing generic attributes record how the attribute block was separated,
    // so the layout lives in a side table rather than in every paragraph.
    private static readonly ConditionalWeakTable<ParagraphBlock, GenericAttributeLayout> GenericAttributeLayouts = new();

    private sealed class GenericAttributeLayout {
        internal string ConsumedWhitespace = string.Empty;
        internal bool SuppressSeparator;
    }

    internal string GenericAttributeConsumedWhitespace {
        get => GenericAttributeLayouts.TryGetValue(this, out var layout) ? layout.ConsumedWhitespace : string.Empty;
        set {
            if (!string.IsNullOrEmpty(value) || GenericAttributeLayouts.TryGetValue(this, out _)) {
                GenericAttributeLayouts.GetValue(this, static _ => new GenericAttributeLayout()).ConsumedWhitespace = value ?? string.Empty;
            }
        }
    }

    internal bool GenericAttributeSuppressSeparator {
        get => GenericAttributeLayouts.TryGetValue(this, out var layout) && layout.SuppressSeparator;
        set {
            if (value || GenericAttributeLayouts.TryGetValue(this, out _)) {
                GenericAttributeLayouts.GetValue(this, static _ => new GenericAttributeLayout()).SuppressSeparator = value;
            }
        }
    }
    /// <summary>Creates a paragraph block.</summary>
    public ParagraphBlock(InlineSequence inlines) { Inlines = inlines; }
    /// <inheritdoc />
    string IMarkdownBlock.RenderMarkdown() {
        if (Attributes.IsEmpty) {
            return Inlines.RenderMarkdown();
        }

        var separator = GenericAttributeSuppressSeparator
            ? string.Empty
            : string.IsNullOrEmpty(GenericAttributeConsumedWhitespace)
                ? " "
                : GenericAttributeConsumedWhitespace;
        return Inlines.RenderMarkdown() + separator + MarkdownAttributeBlockRenderer.RenderInlineTrailing(Attributes);
    }
    /// <inheritdoc />
    string IMarkdownBlock.RenderHtml() {
        var sb = new StringBuilder();
        AppendHtml(sb);
        return sb.ToString();
    }

    void IHtmlAppendableMarkdownBlock.AppendHtml(StringBuilder builder) => AppendHtml(builder);

    private void AppendHtml(StringBuilder sb) {
        sb.Append("<p").Append(MarkdownHtmlAttributes.Render(Attributes, null)).Append('>');
        Inlines.AppendHtml(sb);
        sb.Append(RenderGenericAttributeConsumedWhitespace()).Append("</p>");
    }
    InlineSequence IParagraphMarkdownBlock.ParagraphInlines => Inlines;
    string ITightListItemHtmlMarkdownBlock.RenderTightListItemHtml() => Inlines.RenderHtml();
    InlineSequence IInlineSyntaxMarkdownBlock.SyntaxInlines => Inlines;
    MarkdownSyntaxKind IInlineSyntaxMarkdownBlock.SyntaxKind => MarkdownSyntaxKind.Paragraph;
    MarkdownSourceSpan? IInlineSyntaxMarkdownBlock.ProvidedSyntaxSpan => null;
    MarkdownSyntaxNode ISyntaxMarkdownBlock.BuildSyntaxNode(MarkdownSourceSpan? span) =>
        MarkdownBlockSyntaxBuilder.BuildInlineBlock(this, span);

    private string RenderGenericAttributeConsumedWhitespace() {
        if (Attributes.IsEmpty) {
            return string.Empty;
        }

        string consumedWhitespace = GenericAttributeConsumedWhitespace;
        return consumedWhitespace.Length == 0
            ? string.Empty
            : HtmlTextEncoder.Encode(consumedWhitespace, HtmlRenderContext.Options);
    }
}
