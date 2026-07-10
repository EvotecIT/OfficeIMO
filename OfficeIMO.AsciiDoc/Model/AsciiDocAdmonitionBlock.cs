namespace OfficeIMO.AsciiDoc;

/// <summary>Built-in AsciiDoc admonition kinds.</summary>
public enum AsciiDocAdmonitionKind {
    /// <summary>Note.</summary>
    Note = 0,
    /// <summary>Tip.</summary>
    Tip,
    /// <summary>Important information.</summary>
    Important,
    /// <summary>Warning.</summary>
    Warning,
    /// <summary>Caution.</summary>
    Caution
}

/// <summary>Source-backed admonition paragraph.</summary>
public sealed class AsciiDocAdmonitionBlock : AsciiDocBlock {
    private string _text;
    private bool _textWasAssigned;

    internal AsciiDocAdmonitionBlock(
        AsciiDocSyntaxNode syntax,
        AsciiDocAdmonitionKind kind,
        string label,
        string text,
        AsciiDocInlineSequence inlines,
        string trailingLineEnding) : base(syntax, trailingLineEnding) {
        Kind = kind;
        Label = label;
        _text = text;
        Inlines = inlines;
    }

    /// <summary>Admonition kind.</summary>
    public AsciiDocAdmonitionKind Kind { get; }

    /// <summary>Original uppercase label.</summary>
    public string Label { get; }

    /// <summary>Admonition content after the label.</summary>
    public string Text {
        get => !_textWasAssigned && Inlines.IsModified ? Inlines.ToAsciiDoc() : _text;
        set {
            string normalized = value ?? string.Empty;
            AsciiDocText.EnsureSingleLine(normalized, nameof(value));
            if (SetValue(ref _text, normalized)) _textWasAssigned = true;
        }
    }

    /// <summary>Typed inline content.</summary>
    public AsciiDocInlineSequence Inlines { get; }

    /// <inheritdoc />
    public override bool IsModified => base.IsModified || Inlines.IsModified;

    internal override string WriteCore(AsciiDocWriterContext context) =>
        Label + ":" + ((_textWasAssigned ? _text : Inlines.Write(context)).Length == 0
            ? string.Empty
            : " " + (_textWasAssigned ? _text : Inlines.Write(context))) + EffectiveTrailingLineEnding(context);
}
