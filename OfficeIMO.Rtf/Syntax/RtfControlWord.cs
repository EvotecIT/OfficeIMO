namespace OfficeIMO.Rtf.Syntax;

/// <summary>
/// Represents an RTF control word in a syntax tree.
/// </summary>
public sealed class RtfControlWord : RtfNode {
    internal RtfControlWord(int position, string name, int? parameter, bool hasParameter, string rawText)
        : base(position) {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        Parameter = parameter;
        HasParameter = hasParameter;
        RawText = rawText ?? string.Empty;
    }

    /// <summary>Control word name without the leading backslash.</summary>
    public string Name { get; }

    /// <summary>Optional numeric parameter.</summary>
    public int? Parameter { get; }

    /// <summary>Whether a parameter was explicitly supplied.</summary>
    public bool HasParameter { get; }

    /// <summary>Raw source text.</summary>
    public string RawText { get; }
}
