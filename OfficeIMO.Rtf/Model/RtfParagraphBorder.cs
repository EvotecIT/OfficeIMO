namespace OfficeIMO.Rtf;

/// <summary>
/// Border formatting for one side of an RTF paragraph.
/// </summary>
public sealed class RtfParagraphBorder {
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
}
