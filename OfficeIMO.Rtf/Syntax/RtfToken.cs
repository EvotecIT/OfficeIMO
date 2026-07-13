namespace OfficeIMO.Rtf.Syntax;

/// <summary>
/// Represents one lexical token from an RTF stream.
/// </summary>
public sealed class RtfToken {
    private readonly byte[]? _binaryData;

    internal RtfToken(
        RtfTokenKind kind,
        int position,
        string rawText,
        string? text = null,
        string? controlName = null,
        char? controlSymbol = null,
        int? parameter = null,
        bool hasParameter = false,
        byte[]? binaryData = null) {
        Kind = kind;
        Position = position;
        RawText = rawText ?? string.Empty;
        Text = text;
        ControlName = controlName;
        ControlSymbol = controlSymbol;
        Parameter = parameter;
        HasParameter = hasParameter;
        _binaryData = binaryData == null ? null : (byte[])binaryData.Clone();
    }

    /// <summary>Token kind.</summary>
    public RtfTokenKind Kind { get; }

    /// <summary>Zero-based input position.</summary>
    public int Position { get; }

    /// <summary>Raw source text that produced the token.</summary>
    public string RawText { get; }

    /// <summary>Literal text value for <see cref="RtfTokenKind.Text"/> tokens.</summary>
    public string? Text { get; }

    /// <summary>Control word name without the leading backslash.</summary>
    public string? ControlName { get; }

    /// <summary>Control symbol character without the leading backslash.</summary>
    public char? ControlSymbol { get; }

    /// <summary>Optional numeric parameter for control words and hex-encoded symbols.</summary>
    public int? Parameter { get; }

    /// <summary>Whether the control explicitly supplied a numeric parameter.</summary>
    public bool HasParameter { get; }

    /// <summary>Binary payload following a <c>\binN</c> control word.</summary>
    public byte[]? BinaryData => _binaryData == null ? null : (byte[])_binaryData.Clone();

    /// <inheritdoc />
    public override string ToString() {
        return Kind switch {
            RtfTokenKind.ControlWord => HasParameter ? "\\" + ControlName + Parameter?.ToString(CultureInfo.InvariantCulture) : "\\" + ControlName,
            RtfTokenKind.ControlSymbol => "\\" + ControlSymbol,
            RtfTokenKind.Text => Text ?? string.Empty,
            _ => Kind.ToString()
        };
    }
}
