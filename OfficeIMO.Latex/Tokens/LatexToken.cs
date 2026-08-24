namespace OfficeIMO.Latex;

/// <summary>Lossless TeX-aware token kind.</summary>
public enum LatexTokenKind {
    /// <summary>Ordinary source text.</summary>
    Text = 0,
    /// <summary>Control word or control symbol beginning with backslash.</summary>
    Command,
    /// <summary>Opening required-group brace.</summary>
    OpenBrace,
    /// <summary>Closing required-group brace.</summary>
    CloseBrace,
    /// <summary>Opening optional-argument bracket.</summary>
    OpenBracket,
    /// <summary>Closing optional-argument bracket.</summary>
    CloseBracket,
    /// <summary>Single or double dollar math shift.</summary>
    MathShift,
    /// <summary>Comment from percent marker through end of line content.</summary>
    Comment,
    /// <summary>Spaces and tabs.</summary>
    Whitespace,
    /// <summary>LF, CRLF, or CR token.</summary>
    LineEnding,
    /// <summary>Alignment tab.</summary>
    AlignmentTab,
    /// <summary>Superscript marker.</summary>
    Superscript,
    /// <summary>Subscript marker.</summary>
    Subscript,
    /// <summary>Parameter marker.</summary>
    Parameter,
    /// <summary>Non-breaking-space marker.</summary>
    NonBreakingSpace,
    /// <summary>Opaque inline verbatim command or verbatim-like environment.</summary>
    Verbatim
}

/// <summary>Exact token from decoded LaTeX source.</summary>
public sealed class LatexToken {
    private readonly LatexSourceText _source;
    private readonly int _startOffset;
    private readonly int _endOffset;
    private string? _text;

    internal LatexToken(
        LatexTokenKind kind,
        LatexSourceText source,
        string? value,
        int startOffset,
        int endOffset,
        bool isTerminated = true) {
        Kind = kind;
        _source = source;
        _startOffset = startOffset;
        _endOffset = endOffset;
        Value = value;
        IsTerminated = isTerminated;
    }

    /// <summary>Token kind.</summary>
    public LatexTokenKind Kind { get; }
    /// <summary>Exact token source.</summary>
    public string Text => _text ??= _source.Text.Substring(_startOffset, _endOffset - _startOffset);
    /// <summary>Command name without backslash, or null.</summary>
    public string? Value { get; }
    /// <summary>Exact source span.</summary>
    public LatexSourceSpan Span => _source.CreateSpan(_startOffset, _endOffset);
    /// <summary>Whether an opaque verbatim region contained its expected closing delimiter.</summary>
    public bool IsTerminated { get; }

    internal int StartOffset => _startOffset;
    internal int EndOffset => _endOffset;
}
