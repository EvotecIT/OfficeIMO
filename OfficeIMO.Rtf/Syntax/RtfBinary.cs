namespace OfficeIMO.Rtf.Syntax;

/// <summary>
/// Represents raw binary payload from a <c>\binN</c> control word.
/// </summary>
public sealed class RtfBinary : RtfNode {
    private readonly byte[]? _data;

    internal RtfBinary(int position, byte[]? data, string rawText, bool deferred = false, int? declaredLength = null)
        : base(position) {
        _data = deferred ? null : data == null ? Array.Empty<byte>() : (byte[])data.Clone();
        RawText = rawText ?? string.Empty;
        DeclaredLength = declaredLength ?? Length;
    }

    /// <summary>Binary payload bytes.</summary>
    public byte[] Data => _data == null ? DecodeRawText(RawText) : (byte[])_data.Clone();
    /// <summary>Available source bytes without materializing a deferred payload.</summary>
    public int Length => _data?.Length ?? RawText.Length;
    internal int DeclaredLength { get; }
    internal static byte[] DecodeRawText(string rawText) {
        var bytes = new byte[rawText.Length];
        for (int index = 0; index < bytes.Length; index++) bytes[index] = (byte)(rawText[index] & 255);
        return bytes;
    }

    /// <summary>Raw binary payload as it appeared in the source text.</summary>
    public string RawText { get; }
}
