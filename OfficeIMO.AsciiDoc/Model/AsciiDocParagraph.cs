namespace OfficeIMO.AsciiDoc;

/// <summary>AsciiDoc paragraph with source-preserving block boundaries.</summary>
public sealed class AsciiDocParagraph : AsciiDocBlock {
    private readonly AsciiDocEditableInlineContent _content;

    internal AsciiDocParagraph(AsciiDocSyntaxNode syntax, string text, AsciiDocInlineSequence inlines, string trailingLineEnding)
        : base(syntax, trailingLineEnding) {
        _content = new AsciiDocEditableInlineContent(text, inlines);
    }

    /// <summary>Paragraph text with line endings normalized to line feeds.</summary>
    public string Text {
        get => AsciiDocText.NormalizeLineEndings(_content.Text, "\n");
        set {
            string normalized = AsciiDocText.NormalizeLineEndings(value ?? string.Empty, "\n");
            _content.Assign(normalized);
        }
    }

    /// <summary>Typed lossless inline content in the paragraph.</summary>
    public AsciiDocInlineSequence Inlines => _content.Inlines;

    /// <inheritdoc />
    public override bool IsModified => base.IsModified || _content.IsModified;

    internal override string WriteCore(AsciiDocWriterContext context) =>
        AsciiDocText.NormalizeLineEndings(AsciiDocLiteralText.EscapeBlockStarts(_content.Write(context)), context.LineEnding) + EffectiveTrailingLineEnding(context);
}
