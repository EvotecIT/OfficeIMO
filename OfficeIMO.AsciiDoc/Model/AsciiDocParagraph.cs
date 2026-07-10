namespace OfficeIMO.AsciiDoc;

/// <summary>AsciiDoc paragraph with source-preserving block boundaries.</summary>
public sealed class AsciiDocParagraph : AsciiDocBlock {
    private string _text;
    private bool _textWasAssigned;

    internal AsciiDocParagraph(AsciiDocSyntaxNode syntax, string text, AsciiDocInlineSequence inlines, string trailingLineEnding)
        : base(syntax, trailingLineEnding) {
        _text = text;
        Inlines = inlines;
    }

    /// <summary>Paragraph text with line endings normalized to line feeds.</summary>
    public string Text {
        get => !_textWasAssigned && Inlines.IsModified
            ? AsciiDocText.NormalizeLineEndings(Inlines.ToAsciiDoc(), "\n")
            : _text;
        set {
            string normalized = AsciiDocText.NormalizeLineEndings(value ?? string.Empty, "\n");
            if (SetValue(ref _text, normalized)) _textWasAssigned = true;
        }
    }

    /// <summary>Typed lossless inline content in the paragraph.</summary>
    public AsciiDocInlineSequence Inlines { get; }

    /// <inheritdoc />
    public override bool IsModified => base.IsModified || Inlines.IsModified;

    internal override string WriteCore(AsciiDocWriterContext context) =>
        (_textWasAssigned
            ? AsciiDocText.NormalizeLineEndings(_text, context.LineEnding)
            : Inlines.Write(context)) + EffectiveTrailingLineEnding(context);
}
