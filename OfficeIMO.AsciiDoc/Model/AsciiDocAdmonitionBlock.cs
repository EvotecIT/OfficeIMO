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
    private readonly AsciiDocEditableInlineContent _content;

    internal AsciiDocAdmonitionBlock(
        AsciiDocSyntaxNode syntax,
        AsciiDocAdmonitionKind kind,
        string label,
        string text,
        AsciiDocInlineSequence inlines,
        string trailingLineEnding) : base(syntax, trailingLineEnding) {
        Kind = kind;
        Label = label;
        _content = new AsciiDocEditableInlineContent(text, inlines);
    }

    /// <summary>Admonition kind.</summary>
    public AsciiDocAdmonitionKind Kind { get; }

    /// <summary>Original uppercase label.</summary>
    public string Label { get; }

    /// <summary>Admonition content after the label.</summary>
    public string Text {
        get => _content.Text;
        set {
            string normalized = value ?? string.Empty;
            AsciiDocText.EnsureSingleLine(normalized, nameof(value));
            _content.Assign(normalized);
        }
    }

    /// <summary>Typed inline content.</summary>
    public AsciiDocInlineSequence Inlines => _content.Inlines;

    /// <inheritdoc />
    public override bool IsModified => base.IsModified || _content.IsModified;

    internal override string WriteCore(AsciiDocWriterContext context) {
        string text = _content.Write(context);
        return Label + ":" + (text.Length == 0 ? string.Empty : " " + text) + EffectiveTrailingLineEnding(context);
    }
}
