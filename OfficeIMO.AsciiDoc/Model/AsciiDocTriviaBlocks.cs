namespace OfficeIMO.AsciiDoc;

/// <summary>Source-preserved blank line.</summary>
public sealed class AsciiDocBlankLine : AsciiDocBlock {
    internal AsciiDocBlankLine(AsciiDocSyntaxNode syntax, string lineEnding) : base(syntax, lineEnding) { }

    internal override string WriteCore(AsciiDocWriterContext context) =>
        context.Mode == AsciiDocWriterMode.Preserve ? OriginalText : context.LineEnding;
}

/// <summary>Single-line AsciiDoc comment.</summary>
public sealed class AsciiDocLineComment : AsciiDocBlock {
    private string _text;

    internal AsciiDocLineComment(AsciiDocSyntaxNode syntax, string text, string trailingLineEnding)
        : base(syntax, trailingLineEnding) {
        _text = text;
    }

    /// <summary>Comment text without the leading <c>//</c> marker.</summary>
    public string Text {
        get => _text;
        set {
            string normalized = value ?? string.Empty;
            AsciiDocText.EnsureSingleLine(normalized, nameof(value));
            SetValue(ref _text, normalized);
        }
    }

    internal override string WriteCore(AsciiDocWriterContext context) =>
        "//" + (Text.Length == 0 ? string.Empty : " " + Text) + EffectiveTrailingLineEnding(context);
}

/// <summary>Source-preserved block not interpreted by the current semantic profile.</summary>
public sealed class AsciiDocRawBlock : AsciiDocBlock {
    internal AsciiDocRawBlock(AsciiDocSyntaxNode syntax, string trailingLineEnding) : base(syntax, trailingLineEnding) { }

    internal override string WriteCore(AsciiDocWriterContext context) => OriginalText;
}
