namespace OfficeIMO.Rtf;

/// <summary>
/// Embedded font metadata stored inside a font table entry.
/// </summary>
public sealed class RtfFontEmbedding {
    /// <summary>Embedded font type.</summary>
    public RtfEmbeddedFontType Type { get; set; }

    /// <summary>Optional file name from the <c>{\*\fontfile ...}</c> destination.</summary>
    public string? FileName { get; set; }

    /// <summary>Optional file name code page from <c>\cpg</c> inside <c>\fontfile</c>.</summary>
    public int? FileCodePage { get; set; }

    /// <summary>Embedded font payload bytes.</summary>
    public byte[] Data { get; set; } = Array.Empty<byte>();
}
