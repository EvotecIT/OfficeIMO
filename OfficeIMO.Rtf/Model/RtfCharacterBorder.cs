namespace OfficeIMO.Rtf;

/// <summary>
/// Character border formatting for an RTF text run.
/// </summary>
public sealed class RtfCharacterBorder {
    /// <summary>Border line style.</summary>
    public RtfParagraphBorderStyle Style { get; set; } = RtfParagraphBorderStyle.None;

    /// <summary>Border width value carried by the RTF <c>\brdrw</c> control.</summary>
    public int? Width { get; set; }

    /// <summary>One-based color table index.</summary>
    public int? ColorIndex { get; set; }

    /// <summary>Whether any border formatting is present.</summary>
    public bool HasAnyValue =>
        Style != RtfParagraphBorderStyle.None ||
        Width.HasValue ||
        ColorIndex.HasValue;

    internal void Clear() {
        Style = RtfParagraphBorderStyle.None;
        Width = null;
        ColorIndex = null;
    }

    internal void CopyFrom(RtfCharacterBorder source) {
        Style = source.Style;
        Width = source.Width;
        ColorIndex = source.ColorIndex;
    }
}
