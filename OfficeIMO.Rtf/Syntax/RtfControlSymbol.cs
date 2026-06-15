namespace OfficeIMO.Rtf.Syntax;

/// <summary>
/// Represents an RTF control symbol in a syntax tree.
/// </summary>
public sealed class RtfControlSymbol : RtfNode {
    internal RtfControlSymbol(int position, char symbol, int? parameter, bool hasParameter, string rawText)
        : base(position) {
        Symbol = symbol;
        Parameter = parameter;
        HasParameter = hasParameter;
        RawText = rawText ?? string.Empty;
    }

    /// <summary>Symbol character without the leading backslash.</summary>
    public char Symbol { get; }

    /// <summary>Optional numeric parameter, used for hex-encoded symbols.</summary>
    public int? Parameter { get; }

    /// <summary>Whether a parameter was explicitly supplied.</summary>
    public bool HasParameter { get; }

    /// <summary>Raw source text.</summary>
    public string RawText { get; }
}
