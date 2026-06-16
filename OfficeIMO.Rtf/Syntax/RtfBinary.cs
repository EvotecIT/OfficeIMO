namespace OfficeIMO.Rtf.Syntax;

/// <summary>
/// Represents raw binary payload from a <c>\binN</c> control word.
/// </summary>
public sealed class RtfBinary : RtfNode {
    internal RtfBinary(int position, byte[] data, string rawText)
        : base(position) {
        Data = data ?? Array.Empty<byte>();
        RawText = rawText ?? string.Empty;
    }

    /// <summary>Binary payload bytes.</summary>
    public byte[] Data { get; }

    /// <summary>Raw binary payload as it appeared in the source text.</summary>
    public string RawText { get; }
}
