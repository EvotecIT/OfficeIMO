namespace OfficeIMO.Rtf.Syntax;

/// <summary>
/// Represents raw binary payload from a <c>\binN</c> control word.
/// </summary>
public sealed class RtfBinary : RtfNode {
    private readonly byte[] _data;

    internal RtfBinary(int position, byte[] data, string rawText)
        : base(position) {
        _data = data == null ? Array.Empty<byte>() : (byte[])data.Clone();
        RawText = rawText ?? string.Empty;
    }

    /// <summary>Binary payload bytes.</summary>
    public byte[] Data => (byte[])_data.Clone();

    /// <summary>Raw binary payload as it appeared in the source text.</summary>
    public string RawText { get; }
}
